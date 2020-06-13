# 001

- Quando for gravar uma macro:
"Armazenar macro em" -> Pasta de trabalho pessoal de macros
	- permite usar em várias pastas
	- use o campo de descrição para documentar

- Pratique VBA para não esquecer, além de testar as diferentes possibilidades e presenciar os erros
- Uma macro armazena o zoom da planilha, por exemplo
- Lembre-se de salvar o arquivo como pasta de trabalho havilitada para macros do Excel.
________________________________________________
# 002

- CurrentRegion
	- Equivale ao atalho Ctrl + *, seleciona a planilha em questão.

- Pasta "Personal" é um arquivo oculto
	- Entre em exibir -> reexibir -> selecionar "Personal"
	- Não altere a planilha Personal!
	- Salvar a planilha do usuário e fechar. Ir na planilha Personal e ocultá-la novamente. Ao fechar, salve as alterações! (da planilha personal)

- Observações
	- Ctrl + Enter em uma seleção cola o valor da primeira célula em todas as outras
			x   x   x   x   x
			x   x   x   x   x
	- Comando Let é opcional e poucas pessoas usam. Seria usado em uma variável que recebe outro valor.
		>>> Let LastRow = Rows.Count
	- Verificação imediata (abre um console): Ctrl + G

DIM
- Dimensionar uma variável significa avisar ao VBA o espaço máximo que ela vai ocupar. A consequência disso é que ele separa somente a memória necessária para armazenar seus valores, o que é bom por salvar memória, mas pode dar erro se o programador tentar guardar um valor além do range.
- A tabela de tipos de variáveis, seus tamanhos e os ranges estão no site da Microsoft
- O que é importante saber está na próxima aula

SET
- Uma variável só recebe um OBJETO (Workbook, Worksheet) com o comando Set
>>> Dim NewSheet As WorkBook
    Set NewSheet = ActiveSheet

' Parte do código:
>>> Sub Aula2_ConfigurarPágina()
    Dim Celulas As Range
    Range("B3").Select
    Set Celulas = Selection.CurrentRegion
    ActiveSheet.PageSetup.PrintArea = Celulas.Adress

____________________________________________________
# 003

- Referências relativas e Offset

Offset([posiçãoLinha], [posiçãoColuna])
negativo:      sobe        esquerda
positivo:      desce       direita

>>> ActiveCell.Offset(2, 2).Select
    ActiveCell.Offset(-1, 0).Select
    ActiveCell.Offset(1, 0).Select