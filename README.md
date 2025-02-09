Bot-Handler-PDF

Um bot automatizado para manipulação de faturas em PDF. Este bot é capaz de dividir páginas, renomear arquivos com base em informações extraídas de seu conteúdo e gerar tabelas XLSX para organização e controle.

Funcionalidades:

Divisão de PDFs extensos: Divide arquivos PDF com múltiplas páginas em arquivos individuais de uma página.
Renomeação inteligente: Renomeia arquivos PDF automaticamente com base no nome do pagador extraído do arquivo.
Criação de tabelas XLSX: Gera tabelas organizadas contendo informações como:

Nome do pagador.
Data de vencimento.
Mensagem personalizada.
Localização do arquivo renomeado.

Suporte a múltiplos bancos: Atualmente suporta PDFs gerados pelo Asaas e Sicoob. Suporte para outros bancos pode ser adicionado no futuro.
Mensagem personalizada: Permite que o usuário insira uma mensagem personalizada, como:

"Obrigado {cliente} pela assinatura"

Onde {cliente} é substituído pelo nome do pagador correspondente.

Interface intuitiva: Uma interface gráfica simples e funcional que:

Permite selecionar pastas de entrada e saída.
Mostra uma tabela em tempo real com os arquivos processados.
Exibe logs detalhados do processo.
Oferece a opção de exportar os dados para um arquivo XLSX.

Como funciona?

O usuário seleciona uma pasta de entrada contendo PDFs de boletos ou faturas.
O bot analisa os arquivos, dividindo PDFs com múltiplas páginas e extraindo informações relevantes, como:

Nome do pagador.
Data de vencimento.

Os arquivos são renomeados com base no nome do pagador e movidos para a pasta de saída escolhida.
As informações extraídas são organizadas em uma tabela visual na interface, permitindo fácil visualização e controle.
O usuário pode exportar os dados processados para um arquivo XLSX para uso em ferramentas como Microsoft Excel.
Pré-requisitos
Certifique-se de ter as seguintes bibliotecas instaladas antes de executar o bot:

os
shutil
PyPDF2
pdfplumber
tkinter
threading
pandas

Para instalar as dependências necessárias, execute:

pip install PyPDF2 pdfplumber pandas

Como usar

Execute o script principal:

python main.py
Na interface:

Selecione a pasta de entrada com os PDFs a serem processados.
Selecione a pasta de saída onde os arquivos renomeados serão armazenados.
(Opcional) Insira uma mensagem personalizada que será incluída na tabela.
Clique em Iniciar Organização.
Após o processamento, exporte os resultados para um arquivo XLSX, se necessário.

Exemplo de uso
Imagine que você tenha uma pasta contendo boletos em PDF com múltiplas páginas. Após processá-los com o bot:

Os arquivos serão divididos, caso necessário.
Cada arquivo será renomeado com o nome do pagador.
Uma tabela será criada com informações organizadas, como data de vencimento e mensagens personalizadas.

(Arquivos com Nomes iguais nao sao repetidos, caso tenha 2 ou + arquivos com o mesmo nome, apenas um sera movido para pasta, porem, na tabela XLSX, sera citado todos os arquivos, entao por la, o usuario podera verificar caso haja discrepancia na quantidade de arquivos finais gerados, e o total de arquivos manipulados pelo BOT).


