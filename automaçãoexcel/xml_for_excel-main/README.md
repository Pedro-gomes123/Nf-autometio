
# 🧾 XML para Excel - Notas Fiscais Eletrônicas (NF-e)

Este projeto em Python faz a leitura de arquivos XML de notas fiscais eletrônicas (NF-e) e exporta as informações relevantes para uma planilha Excel de forma estruturada.

---

## ✅ Funcionalidades

- Extração de dados das notas fiscais no formato XML.
- Transformação dos dados em uma planilha `.xlsx`.
- Separação dos campos de endereço em colunas distintas: logradouro, número, bairro, município, UF e CEP.
- Detecção automática do peso informado na nota (se disponível).

---

## 📂 Estrutura de Pastas

```
projeto/
│
├── nfs/                 # Pasta onde os arquivos XML devem ser colocados
│   ├── arquivo1.xml
│   ├── arquivo2.xml
│   ├── arquivo3.xml
├── main.py       # Script principal
├── NotasFiscais.xlsx    # Arquivo Excel gerado (após rodar o script)
└── README.md
```

---

## ▶️ Como usar

1. Coloque todos os arquivos `.xml` das notas fiscais na pasta `nfs/`.
2. Execute o script `main.py`:
   ```bash
   python main.py
   ```
3. O Excel `NotasFiscais.xlsx` será criado na raiz do projeto.

---

## 📌 Dependências

Instale os pacotes com:

```bash
pip install xmltodict pandas openpyxl
```

---

## 🛠️ Campos Exportados

| Campo               | Descrição                           |
|---------------------|-------------------------------------|
| numero_nota         | ID da nota                          |
| empresa_emissora    | Nome da empresa que emitiu a nota   |
| nome_cliente        | Nome do destinatário                |
| logradouro          | Rua ou avenida                      |
| numero              | Número do endereço                  |
| bairro              | Bairro                              |
| municipio           | Cidade                              |
| uf                  | Unidade Federativa (Estado)         |
| cep                 | Código postal                       |
| peso                | Peso bruto informado na nota        |

---

## 🧠 Personalizações

Você pode facilmente adicionar ou remover campos do XML conforme a necessidade da empresa ou cliente. Também é possível integrar com bancos de dados ou enviar o Excel por e-mail automaticamente.

---

## ✨ Exemplo de uso

Ideal para:

- Escritórios contábeis que recebem XMLs dos clientes.
- Empreendedores que desejam automatizar a conferência de NF-es.
- Profissionais de dados que querem estruturar informações fiscais para análise.

---

## 📫 Contato

Desenvolvido por Raphael Macedo  
📧 raphael.macedo.oliveira@hotmail.com  
🔗 [LinkedIn - Raphael Macedo](https://www.linkedin.com/in/raphael-macedo10/)

---
