# Extrator-PDF

Programa feito durante estágio na STRC/PA, locada na Unidade Regional da Polícia Civil de Minas Gerais/Pouso Alegre
Feito como presente de despedida

Foi minha primeira experiência fazendo algo que realmente tem uma função prática, no caso, geração de etiquetas e cadastro de materiais na Unidade Regional de Custódia (URC).

O intuito foi agilizar todo o processo de Identificação e Constatação de Drogas de Abuso, onde os perítos perdiam cerca de 2 horas ou mais para a análise de cerca de 15 drogas,
onde com a ajuda do programa, esse tempo foi reduzido para aproximadamente 40 minutos. Uma redução de 67% do tempo necessário para o trabalho.
Além disso, o aplicativo é também utilizado no cadastro de todo o material que sai dos perítos que chegam na URC, sendo uma única pessoa responsável por toda a carga vinda de 8 perítos, o que aumenta muito a necessidade da agilidade e precisão do processo.

O programa coleta dados do PDF gerado pelo próprio serviço padrão da polícia civil, o PCnet, e trata os dados, e os transfere para um arquivo Excel, onde com a utilização de uma impressora etiquetadora e do Software Bartender da Seagull, esses dados são utilizados para a impressão de etiquetas de identificação. Da mesma forma, os dados são utilizados para o cadastro desses materiais na URC de forma que cada Ficha de Acompanhamento de Vestígio (FAV) que é um número único, consiga obter todo o restante dos dados de forma automática, diminuindo o tempo necessário de digitação e também evita erros, uma vez que todo dado é capturado automaticamente.

Para a sua utilização basta abri-lo, selecionar arquivo .PDF que contém os dados a serem capturados e o programa faz o resto, os salvando em um arquivo excel que pode ser aberto para conferência ou edição de algo que seja necessário.

O programa faz o uso das bibliotecas Pandas, Regex, Tkinter e PyPDF2, cada um com uma finalidade dentro de todo o processo.
