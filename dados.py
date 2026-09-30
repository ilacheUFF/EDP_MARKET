# dados.py
# Base dos 66 Eco-design Principles (EDPs)
# Nomes conforme Maccioni, Borgianni & Pigosso (2019), Tabela 2.
# 'etapa' é uma classificação operacional para o jogo EDP MARKET.

ORCAMENTO_INICIAL = 1000
ETAPAS = ["Extração / Materiais", "Produção", "Distribuição", "Uso", "Fim de vida"]

EDPS = [
    {"codigo": "P01", "nome": "Design for appropriate lifespan", "etapa": "Uso"},
    {"codigo": "P02", "nome": "Design for reliability (simplify, reduce number of component)", "etapa": "Uso"},
    {"codigo": "P03", "nome": "Facilitate upgrading and adaptability", "etapa": "Uso"},
    {"codigo": "P04", "nome": "Facilitate maintenance", "etapa": "Uso"},
    {"codigo": "P05", "nome": "Facilitate repair", "etapa": "Uso"},
    {"codigo": "P06", "nome": "Facilitate reuse", "etapa": "Fim de vida"},
    {"codigo": "P07", "nome": "Facilitate remanufacturing", "etapa": "Fim de vida"},
    {"codigo": "P08", "nome": "Intensify use (share, multifunction, integrated, on demand)", "etapa": "Uso"},
    {"codigo": "P09", "nome": "Minimize material content or dematerialize, digitalize, miniaturize", "etapa": "Extração / Materiais"},
    {"codigo": "P10", "nome": "Minimize scraps and discards", "etapa": "Produção"},
    {"codigo": "P11", "nome": "Minimize packaging (avoid, integrate, drastically reduce)", "etapa": "Produção"},
    {"codigo": "P12", "nome": "Minimize material consumption during usage (select more consumption-efficient systems)", "etapa": "Uso"},
    {"codigo": "P13", "nome": "Minimize materials consumption during usage (engage systems with dynamic material consumption)", "etapa": "Uso"},
    {"codigo": "P14", "nome": "Minimize energy consumption (during pre-production and production)", "etapa": "Produção"},
    {"codigo": "P15", "nome": "Minimize energy consumption (during transportation and storage)", "etapa": "Distribuição"},
    {"codigo": "P16", "nome": "Minimize energy consumption (select systems with energy-efficient operation stage)", "etapa": "Uso"},
    {"codigo": "P17", "nome": "Select non-toxic and harmless materials", "etapa": "Extração / Materiais"},
    {"codigo": "P18", "nome": "Select non-toxic and harmless energy resources", "etapa": "Produção"},
    {"codigo": "P19", "nome": "Select renewable and bio-compatible materials", "etapa": "Extração / Materiais"},
    {"codigo": "P20", "nome": "Select renewable and bio-compatible energy resources", "etapa": "Produção"},
    {"codigo": "P21", "nome": "Adopt the cascade approach", "etapa": "Fim de vida"},
    {"codigo": "P22", "nome": "Facilitate end-of-life collection and transportation", "etapa": "Fim de vida"},
    {"codigo": "P23", "nome": "Identify materials", "etapa": "Fim de vida"},
    {"codigo": "P24", "nome": "Minimize the overall number of different incompatible materials", "etapa": "Extração / Materiais"},
    {"codigo": "P25", "nome": "Facilitate cleaning", "etapa": "Uso"},
    {"codigo": "P26", "nome": "Facilitate composting", "etapa": "Fim de vida"},
    {"codigo": "P27", "nome": "Reduce and facilitate operations of disassembly and separation", "etapa": "Fim de vida"},
    {"codigo": "P28", "nome": "Engage reversible joining systems", "etapa": "Produção"},
    {"codigo": "P29", "nome": "Develop services providing added value to the product’s life cycle", "etapa": "Uso"},
    {"codigo": "P30", "nome": "Develop services providing ‘final results’", "etapa": "Uso"},
    {"codigo": "P31", "nome": "Develop services providing ‘enabling platforms for customers’", "etapa": "Uso"},
    {"codigo": "P32", "nome": "Change raw material (low impact or recyclable/recycled)", "etapa": "Extração / Materiais"},
    {"codigo": "P33", "nome": "Modify raw material (reduce mass and volume)", "etapa": "Extração / Materiais"},
    {"codigo": "P34", "nome": "Use pre-manufactured material", "etapa": "Produção"},
    {"codigo": "P35", "nome": "Change material (low impact) for the packaging", "etapa": "Produção"},
    {"codigo": "P36", "nome": "Reduce the mass of the packaging (eliminate, reduce, integrate, modify raw material to be protected)", "etapa": "Produção"},
    {"codigo": "P37", "nome": "Reduce the volume of the packaging (reduce the volume, dynamize)", "etapa": "Produção"},
    {"codigo": "P38", "nome": "Improve packaging durability (reuse in the same or different way, recycle)", "etapa": "Distribuição"},
    {"codigo": "P39", "nome": "Modify packaging to increase raw material durability (controlled atmosphere, vacuum, deep freeze, sterilization)", "etapa": "Distribuição"},
    {"codigo": "P40", "nome": "Use efficient transports", "etapa": "Distribuição"},
    {"codigo": "P41", "nome": "Optimize logistic (choose the shortest route in time or in space)", "etapa": "Distribuição"},
    {"codigo": "P42", "nome": "Reduce the volume of the load (reduce the volume or the free space)", "etapa": "Distribuição"},
    {"codigo": "P43", "nome": "Reduce the transportation risk", "etapa": "Distribuição"},
    {"codigo": "P44", "nome": "Modify the products for maintenance", "etapa": "Uso"},
    {"codigo": "P45", "nome": "Avoid auxiliary components (independent, eliminate components, dematerializing)", "etapa": "Uso"},
    {"codigo": "P46", "nome": "Reduce auxiliary components (rationalize, improve energy performances, discretize use)", "etapa": "Uso"},
    {"codigo": "P47", "nome": "Remove dangerous materials during the product use", "etapa": "Uso"},
    {"codigo": "P48", "nome": "Reduce environmental problems during the product use", "etapa": "Uso"},
    {"codigo": "P49", "nome": "Act on durability (more operation, different way) of the packaging during the product use", "etapa": "Uso"},
    {"codigo": "P50", "nome": "Reduce the mass of the packaging during the product use (light material, multifunctional product)", "etapa": "Uso"},
    {"codigo": "P51", "nome": "Act on shape and volume", "etapa": "Uso"},
    {"codigo": "P52", "nome": "Reduce the consumption of energy required (efficient or alternative)", "etapa": "Uso"},
    {"codigo": "P53", "nome": "Recover dissipated energy during the product use", "etapa": "Uso"},
    {"codigo": "P54", "nome": "Reduce the need of maintenance during the product use", "etapa": "Uso"},
    {"codigo": "P55", "nome": "Make the maintenance proactive during the product use", "etapa": "Uso"},
    {"codigo": "P56", "nome": "Design repairable product", "etapa": "Uso"},
    {"codigo": "P57", "nome": "Realize modular product", "etapa": "Uso"},
    {"codigo": "P58", "nome": "Design product with low emission of pollution substances during the product use", "etapa": "Uso"},
    {"codigo": "P59", "nome": "Reduce the product mass during the product use", "etapa": "Uso"},
    {"codigo": "P60", "nome": "Reduce the product volume during the product use", "etapa": "Uso"},
    {"codigo": "P61", "nome": "Recovery of material (reuse, remanufacture, recycle) at the end of the product’s life", "etapa": "Fim de vida"},
    {"codigo": "P62", "nome": "Use biodegradable packaging for the end of the product’s life", "etapa": "Fim de vida"},
    {"codigo": "P63", "nome": "Reduce the mass for the end of the product’s life", "etapa": "Fim de vida"},
    {"codigo": "P64", "nome": "Allow the separability of materials at the end of the product’s life", "etapa": "Fim de vida"},
    {"codigo": "P65", "nome": "Use connections between parts for the end of the product’s life", "etapa": "Fim de vida"},
    {"codigo": "P66", "nome": "Reduce the weight and the volume of the transport at the end of the product’s life", "etapa": "Fim de vida"},
]

# Estrutura inicial; preencher depois com os nomes reais da turma.
ALUNOS = ["Lila","Patricia"]

def obter_edp(codigo):
    """Retorna o EDP correspondente ao código."""
    return next((edp for edp in EDPS if edp["codigo"] == codigo), None)

def obter_vendedor(codigo, alunos=None):
    """Retorna o aluno responsável por vender o EDP."""
    if alunos is None:
        alunos = ALUNOS
    return next((aluno for aluno in alunos if aluno.get("edp") == codigo), None)