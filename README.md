# Nf-autometio

🧾 Sistema de Automação de Lançamento de Notas Fiscais
Este projeto é um sistema que automatiza o lançamento de notas fiscais eletrônicas (NFe) a partir de arquivos XML, convertendo e organizando os dados relevantes em uma planilha Excel de forma rápida e eficiente.

🚀 Funcionalidades
📥 Leitura automática de arquivos XML de notas fiscais na pasta nfs.

🔍 Extração de informações específicas dos arquivos XML.

🏷️ Conversão dos dados extraídos em colunas organizadas.

📊 Criação de um DataFrame com pandas.

📤 Geração automática de planilha Excel com as informações processadas.

🛠️ Tecnologias Utilizadas
Python

pandas – manipulação de dados e criação de DataFrames.

xmltodict – conversão de arquivos XML em dicionários.

os – manipulação de arquivos e pastas.

💻 Como Executar o Projeto
Clone o repositório:

bash
Copiar
Editar
git clone https://github.com/seu-usuario/seu-projeto.git
Acesse a pasta do projeto:

bash
Copiar
Editar
cd nome-do-projeto
Instale as dependências:

bash
Copiar
Editar
pip install pandas xmltodict
Coloque os arquivos XML na pasta chamada nfs.

Execute o script:

bash
Copiar
Editar
python main.py
O sistema gerará automaticamente o arquivo tabela_nfs.xlsx com os dados extraídos.

📝 Como Funciona
O sistema percorre todos os arquivos .xml na pasta nfs.

Cada arquivo é lido e transformado em um dicionário.

As informações desejadas são extraídas de locais específicos do XML.

Esses dados são convertidos em colunas, organizados em um DataFrame.

O DataFrame é exportado como uma planilha Excel (.xlsx).

🎯 Vantagens
Economia de tempo com automação de tarefas repetitivas.

Redução de erros humanos no lançamento de dados.

Facilidade na manipulação e análise das informações extraídas.

🤝 Contribuição
Contribuições são bem-vindas!
Sinta-se à vontade para:

Abrir Issues.

Enviar Pull Requests.

Sugerir melhorias.

📄 Licença
Este projeto está sob a licença MIT.

📞 Contato
Desenvolvedor: Pedro de Carvalho Gomes

Email: pedrocarvalhog10@gmail.com
