import unicodedata

def remover_acentos(texto: str) -> str:
    # Normaliza e remove os acentos
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

print(remover_acentos("São Paulo".split('-')[0].strip().upper()) in remover_acentos("Sao Paulo".upper()))