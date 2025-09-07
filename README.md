# SAP OV Automation

Automação em **VBA** e **Python** para otimizar processos relacionados às **Ordens de Venda (OV)** no **SAP**.  
Este projeto integra extração, classificação e validação de dados, além de auxiliar na comunicação com clientes via WhatsApp.

---

## 🚀 Funcionalidades

- **Consultar_Informações**  
  Obtém dados do logradouro do cliente na transação **VA03** e insere automaticamente na tabela `INFORMAÇÕES`.

- **Mensagem_WhatsApp**  
  Gera mensagens personalizadas para clientes com o número da OV.  
  O usuário copia a mensagem com um clique e a envia manualmente pelo WhatsApp.

- **Classificar_E_EncaminharOVs**  
  Automatiza o tratamento das Ordens de Venda (OV) conforme a decisão do colaborador:  

  1. **Logradouro já existe na base**  
     - O colaborador informa o código do logradouro.  
     - O script altera o status da OV.  
     - O código do logradouro é escrito no campo específico da transação **VA02**.  
     - Uma mensagem é registrada para os colaboradores internos dentro da transação.  

  2. **Logradouro não existe e cliente não apresentou comprovação**  
     - O colaborador marca a OV como *Pedido Não Elaborado* na planilha.  
     - O script altera o status da OV.  
     - Escreve uma mensagem para os colaboradores internos justificando a não elaboração.  
     - Gera uma mensagem destinada ao cliente solicitando comprovação.  

  3. **Logradouro não existe e cliente comprovou**  
     - O colaborador cria e informa um novo código de logradouro.  
     - O script cadastra o novo logradouro.  
     - As informações da tabela `INFORMAÇÕES` são transferidas automaticamente para a tabela `Rua Cadastrada`.  

- **Logradouro_Cadastro_Validação**  
  Verifica se o cadastro do logradouro foi processado corretamente no SAP.  
  Em caso positivo, atualiza o status da OV.

- **BuscarTelefoneRapido**  
  Executa o script Python `Extrair_Telefone_v3.py`, que utiliza **Regex** para extrair números de telefone do cliente diretamente na transação **VA03**.

---

## 🛠️ Tecnologias Utilizadas

- **VBA**  
  Automação em planilhas do Excel, integração com SAP GUI e manipulação de dados.

- **Python**  
  Extração de dados utilizando expressões regulares (**Regex**).

---

## 📂 Estrutura do Repositório

sap-ov-automation/
│
├── vba/ # Scripts em VBA
│ ├── Consultar_Informações.bas
│ ├── Mensagem_WhatsApp.bas
│ ├── Classificar_E_EncaminharOVs.bas
│ ├── Logradouro_Cadastro_Validação.bas
│ ├── BuscarTelefoneRapido.bas
│
├── python/ # Scripts em Python
│ └── Extrair_Telefone_v3.py
│
├── examples/ # Exemplos de planilhas prontas
│ └── exemplo_planilha.xlsm
│
│
└── README.md


## ▶️ Como Usar

1. **Clone este repositório:**
   ```bash
   git clone https://github.com/seuusuario/sap-ov-automation.git
Abra o Excel e importe os módulos da pasta /vba.

Configure o Python (se necessário, instale dependências para Regex):

pip install regex
Teste a automação utilizando a planilha em /examples.

🤝 Contribuições

Este repositório tem caráter profissional e de portfólio.
Contribuições não estão abertas neste momento, mas sugestões são bem-vindas via issues.

📌 Observação

Este projeto foi desenvolvido para demonstrar automação de processos no SAP com integração entre VBA e Python.
Todas as informações e exemplos foram anônimizados para preservar dados corporativos.

