import pandas as pd

df = pd.read_excel('relatorio.xlsx')
df = df.dropna(subset=['Descrição', 'Texto do item'])
df = df[df['Descrição'].str.strip() != '']
df = df[df['Texto do item'].str.strip() != '']

# Extrair correia com regex
df['Correia'] = df['Descrição'].str.extract(r'([\dA-Z]+CV\d+)', expand=False).str.strip()

# Sistemas TCLD
sistemas = {
    'TCLD DA 3 BRITAGEM': ['02CV011', '02CV012', '03CV014', '03CV015', '03CV016', '03CV017', '03CV018', '03CV019', '03CV020', '03CV021', '03CV022', '03CV023', '03CV024', '05CV025'],
    'TCLD NORTE': ['02CV001', '02CV006', '02CV007', '02CV008'],
    'PILHA NORTE': ['11CV025'],
    'FAZENDÃO': ['11CV057', '11CV058'],
    'PILHA CENTRO': ['02CV042'],
    'TCLD SUL': ['02CV010'],
    'USINA 3': ['05CV026', '05CV027', '05CV028', '09CV030']
}

correia_to_sistema = {}
for sistema, correias in sistemas.items():
    for correia in correias:
        correia_to_sistema[correia] = sistema

df_filtered = df[df['Correia'].isin(correia_to_sistema.keys())]
df_filtered['Sistema'] = df_filtered['Correia'].map(correia_to_sistema)

print('✅ Linhas com TCLD:', len(df_filtered))
print('\nSistemas encontrados:')
for sistema in df_filtered['Sistema'].unique():
    count = len(df_filtered[df_filtered['Sistema'] == sistema])
    print(f'  - {sistema}: {count} linhas')
