# Coleta de eventos do eSocial utilizando a ferramenta exporta dados do SOC

Trabalho com mais de 500 clientes em uma empresa de segurança do trabalho, fazemos os envios dos eventos de SST do eSocial ao governo
Para isso utilizamos o sistema SOC

Precisavamos gerar:
 - os recibos de envios dos eventos concluidos, 
 - as inconsistencias apresentadas nos envios, 
 - e os envios pendentes 
Isso de uma maneira mais prática e em massa para todos os clientes da nossa base

Para isso fiz uso da ferramenta exporta dados do sistema que nos fornece uma url com uma chave e através dela,
alterando sua string consigo gerar um txt no navegador com os dados que necessitavamos

Então fiz uma automação utilizando selenium, pandas e openpyxl para:
 - ler os codigos e nomes das empresas de uma base em excel
 - coletar os dados no navegador
 - salvar em uma planilha as inconsistencias e pendencias para facilitar as tratativas
 - salvar em uma planilha para cada cliente os recibos de envios concluidos para serem enviados como anexo para as empresas
