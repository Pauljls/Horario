import json
import random
from copy import deepcopy
from pathlib import Path

from main import read_schedule

# ── Parámetros del algoritmo genético ────────────────────────────────────────
POP_SIZE = 200
GENERATIONS = 500
MUTATION_RATE = 0.3
ELITE_SIZE = 10
TOURNAMENT_K = 5

DIAS = ["lunes", "martes", "miercoles", "jueves", "viernes"]


# ── Representación ────────────────────────────────────────────────────────────
# Un individuo es un dict {dia: [materia|None, ...]} con las materias
# permutadas por día. La longitud de cada lista es siempre NUM_BLOQUES.


_TILDES = str.maketrans("áéíóúüñÁÉÍÓÚÜÑ", "aeiouunAEIOUUN")

def _clave_real(dia: str, bloque: dict) -> str | None:
    """Devuelve la clave real del día en el bloque ignorando tildes y encoding roto."""
    for key in bloque:
        limpia = key.lower().translate(_TILDES)
        if limpia == dia:
            return key
    return None


def bloques_por_dia(horario: list[dict]) -> dict[str, list]:
    """Extrae la lista de materias (con None) por día desde el horario leído."""
    resultado = {dia: [] for dia in DIAS}
    for bloque in horario:
        for dia in DIAS:
            clave = _clave_real(dia, bloque)
            resultado[dia].append(bloque.get(clave) if clave else None)
    return resultado


# ── Fitness ───────────────────────────────────────────────────────────────────

def contar_fragmentaciones(materias: list) -> int:
    """
    Cuenta cuántas veces una materia aparece intercalada con otra distinta.
    Ej: [Mat, Com, Mat] → Mat está fragmentada → penalización 1.
    """
    penalizacion = 0
    vistas: dict[str, int] = {}
    ultima = None
    for m in materias:
        if m is None:
            ultima = None
            continue
        if m != ultima:
            if m in vistas:
                # Esta materia ya apareció antes y fue interrumpida
                penalizacion += 1
            vistas[m] = vistas.get(m, 0) + 1
            ultima = m
        else:
            ultima = m
    return penalizacion


def fitness(individuo: dict[str, list]) -> int:
    """Menor es mejor: suma de fragmentaciones en todos los días."""
    return sum(contar_fragmentaciones(individuo[dia]) for dia in DIAS)


# ── Población inicial ─────────────────────────────────────────────────────────

def crear_individuo(base: dict[str, list]) -> dict[str, list]:
    """Permuta aleatoriamente los bloques de cada día."""
    individuo = {}
    for dia in DIAS:
        bloques = base[dia][:]
        random.shuffle(bloques)
        individuo[dia] = bloques
    return individuo


def poblacion_inicial(base: dict[str, list]) -> list[dict[str, list]]:
    return [crear_individuo(base) for _ in range(POP_SIZE)]


# ── Selección por torneo ──────────────────────────────────────────────────────

def torneo(poblacion: list, fitnesses: list[int]) -> dict[str, list]:
    candidatos = random.sample(range(len(poblacion)), TOURNAMENT_K)
    mejor = min(candidatos, key=lambda i: fitnesses[i])
    return deepcopy(poblacion[mejor])


# ── Cruce OX1 (preserva el multiset de elementos) ────────────────────────────

def _ox1(p1: list, p2: list) -> list:
    """Order Crossover: hereda un segmento de p1 y rellena en el orden de p2."""
    n = len(p1)
    a, b = sorted(random.sample(range(n), 2))
    segmento = p1[a:b + 1]
    # Contar cuántas veces aparece cada elemento en el segmento
    conteo = {}
    for x in segmento:
        conteo[x] = conteo.get(x, 0) + 1
    # Tomar de p2 los elementos que aún faltan
    resto = []
    for x in p2:
        if conteo.get(x, 0) > 0:
            conteo[x] -= 1
        else:
            resto.append(x)
    return resto[:a] + segmento + resto[a:]


def cruce(padre1: dict[str, list], padre2: dict[str, list]) -> dict[str, list]:
    return {dia: _ox1(padre1[dia], padre2[dia]) for dia in DIAS}


# ── Mutación (swap de dos bloques en un día) ──────────────────────────────────

def mutar(individuo: dict[str, list]) -> dict[str, list]:
    for dia in DIAS:
        if random.random() < MUTATION_RATE:
            bloques = individuo[dia]
            if len(bloques) >= 2:
                i, j = random.sample(range(len(bloques)), 2)
                bloques[i], bloques[j] = bloques[j], bloques[i]
    return individuo


# ── Algoritmo genético principal ──────────────────────────────────────────────

def algoritmo_genetico(base: dict[str, list]) -> dict[str, list]:
    poblacion = poblacion_inicial(base)
    mejor_global = None
    mejor_fitness = float("inf")

    for gen in range(GENERATIONS):
        fitnesses = [fitness(ind) for ind in poblacion]

        # Actualizar mejor global
        min_idx = min(range(len(fitnesses)), key=lambda i: fitnesses[i])
        if fitnesses[min_idx] < mejor_fitness:
            mejor_fitness = fitnesses[min_idx]
            mejor_global = deepcopy(poblacion[min_idx])
            print(f"  Gen {gen:>4}: fitness = {mejor_fitness}")

        if mejor_fitness == 0:
            break

        # Élite
        elite_idx = sorted(range(len(fitnesses)), key=lambda i: fitnesses[i])[:ELITE_SIZE]
        nueva_pob = [deepcopy(poblacion[i]) for i in elite_idx]

        # Reproducción
        while len(nueva_pob) < POP_SIZE:
            p1 = torneo(poblacion, fitnesses)
            p2 = torneo(poblacion, fitnesses)
            hijo = cruce(p1, p2)
            hijo = mutar(hijo)
            nueva_pob.append(hijo)

        poblacion = nueva_pob

    return mejor_global


# ── Reconstruir horario con las horas originales ──────────────────────────────

def reconstruir_horario(
    horas: list[str], optimizado: dict[str, list]
) -> list[dict]:
    n = len(horas)
    return [
        {
            "hora": horas[i],
            **{dia: optimizado[dia][i] for dia in DIAS},
        }
        for i in range(n)
    ]


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    archivo = Path(__file__).parent / "horarios.xlsx"
    horario = read_schedule(str(archivo), sheet_name="Hoja1")

    horas = [bloque["hora"] for bloque in horario]
    base = bloques_por_dia(horario)

    import pandas as pd

    df_original = pd.DataFrame(horario).rename(columns={"hora": "HORA"}).set_index("HORA")
    df_original.columns = [c.upper() for c in df_original.columns]
    print("Horario original:")
    print(df_original.to_string())
    print(f"\nFitness original: {fitness(base)}")
    print("\nEjecutando algoritmo genético...\n")

    optimizado = algoritmo_genetico(base)

    resultado = reconstruir_horario(horas, optimizado)

    df_opt = pd.DataFrame(resultado).rename(columns={"hora": "HORA"}).set_index("HORA")
    df_opt.columns = [c.upper() for c in df_opt.columns]
    print("\nHorario optimizado:")
    print(df_opt.to_string())
    print(f"\nFitness final: {fitness(optimizado)}")
