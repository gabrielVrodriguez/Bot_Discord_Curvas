import discord
import pandas as pd
from dotenv import dotenv_values

config = dotenv_values(".env")


intents = discord.Intents.default()
intents.message_content = True

client = discord.Client(intents=intents)


class MyClient(discord.Client):
    async def on_ready(self):
        print('Bot está logado! {0}!'.format(self.user))

    @staticmethod
    async def criar_tabela_dinamica(df):
        dados_buscados = [ "Título do anúncio", "# de anúncio", "Preço unitário de venda do anúncio (BRL)", "Unidades"]
        dados_removidos = []
        for coluna in df.columns:
            if coluna not in dados_buscados:
                dados_removidos.append(coluna)
        
        df = df.drop(columns=dados_removidos)

        tabela_dinamica = pd.pivot_table(df, values=["Preço unitário de venda do anúncio (BRL)", "Unidades" ],
                                         index=["# de anúncio", "Título do anúncio"], aggfunc='sum')
        
        tabela_dinamica = tabela_dinamica.rename(columns={"Preço unitário de venda do anúncio (BRL)": "Valor total vendido",
                                                      "Unidades": "Quantidades vendidas"})
        
       # tabela_dinamica = tabela_dinamica[['Quantidade vendida', 'Valor total vendido', '% da participação total']]
        tabela_dinamica.reset_index(inplace=True)
        tabela_dinamica.sort_values(by='Valor total vendido', ascending=False, inplace=True)

        total_vendido_total = tabela_dinamica['Valor total vendido'].sum()
        tabela_dinamica['% participação total'] = round((tabela_dinamica['Valor total vendido'] / total_vendido_total) * 100, 2)
        tabela_dinamica['Valor total vendido'] = round(tabela_dinamica['Valor total vendido'], 2)

        limite_A = total_vendido_total * 0.80
        limite_B = total_vendido_total * 0.95
        limite_C = total_vendido_total
        curva_anuncio = []
        soma_valor_vendido = 0
        for index, row in tabela_dinamica.iterrows():
            soma_valor_vendido += row['Valor total vendido']
            if soma_valor_vendido <= limite_A:
                curva_anuncio.append('A')
            elif soma_valor_vendido <= limite_B:
                curva_anuncio.append('B')
            else:
                curva_anuncio.append('C')
    
        tabela_dinamica['Curva do anúncio'] = curva_anuncio


        return tabela_dinamica

    async def on_message(self, message):
        print('Message from {0.author}: {0.content}'.format(message))
        conteudo = message.content.lower()
        #leitura da mensagem no chat
        if message.author == client.user:
            return
        if conteudo.startswith('/teste'):
            await message.channel.send(f'Juninho está on! QUAL A BOA!')

        elif conteudo.startswith('/curvas') and message.attachments:
            for attachment in message.attachments:
                if attachment.filename.endswith('.xlsx'):
                    await attachment.save(attachment.filename)
                    #leitura do arquivo
                    try:
                        df = pd.read_excel(attachment.filename, skiprows=5)
                        print(df)
                        tabela_dinamica = await self.criar_tabela_dinamica(df)
                        novo_arquivo_excel = "Curvas.xlsx"
                        tabela_dinamica.to_excel(novo_arquivo_excel, index=False)
                    
                        with open(novo_arquivo_excel, 'rb') as file:
                            await message.channel.send("As curvas foram processadas com sucesso!", file=discord.File(file))

                    except Exception as e:
                        await message.channel.send(f"Ocorreu um erro ao processar o arquivo: {e}")

client = MyClient(intents=intents)
client.run(config["BOT_TOKEN"])