==============================
Sistema de Geração de Certidões
==============================

Este sistema automatiza a geração de certidões com histórico do curso com notas e atividades em PDF para alunos de cursos cadastrados no Moodle, com base nos dados da base MySQL do Moodle e modelos Word personalizados.

------------------------------
1. Componentes principais
------------------------------

• busca_id.py
  - Consulta o banco de dados do Moodle e gera arquivos .xlsx com os dados dos alunos de um curso.
  - Pode buscar:
    - Por ID de curso (turma)
    - Por nome do aluno
  - Os arquivos gerados são salvos automaticamente na pasta "Certidoes_Emitidas".

• gerar_certidao.py
  - Lê todas as planilhas .xlsx dentro da pasta "Certidoes_Emitidas".
  - Para cada planilha:
    - Gera uma certidão em PDF para cada aluno usando um modelo Word ("Certidão.docx").
    - Cria uma subpasta com o nome do curso e ID, onde os PDFs e a planilha original são armazenados.
  - As notas das atividades e nota final são buscadas diretamente no banco de dados.
  - A imagem de assinatura ("Assinatura.jpg") é inserida na certidão.

• conn.py
  - Armazena as configurações de conexão com o banco de dados MySQL (host, user, senha, database).
  - Usado por ambos os scripts para evitar exposição direta de credenciais.

------------------------------
2. Estrutura de Pastas Esperada
------------------------------

/
├── Certidoes_Emitidas/
│   ├── alunos_CURSO.xlsx               ← Gerado pelo busca_id.py
│   └── CURSO [id]/                     ← Criado pelo gerar_certidao.py
│       ├── Certidao - ALUNO.pdf
│       └── alunos_CURSO.xlsx           ← Movido para cá após geração
├── Certidao.docx                       ← Modelo do documento Word
├── Assinatura.jpg                      ← Imagem da assinatura
├── busca_id.py
├── gerar_certidao.py
└── conn.py

------------------------------
3. Como usar
------------------------------

ETAPA 1: Gerar a planilha de alunos

- Edite o arquivo "busca_id.py"
  • Para buscar por turma: informe o "id_curso"
  • Para buscar por nome: informe "busca_aluno"
- Execute o script:
    python busca_id.py

ETAPA 2: Gerar as certidões

- Execute:
    python gerar_certidao.py
- O sistema processará automaticamente todos os arquivos .xlsx na pasta "Certidoes_Emitidas".

------------------------------
4. Observações
------------------------------

• A planilha deve conter as colunas: nome_completo, username, curso, id_curso, id_aluno
• O CPF será formatado automaticamente, mesmo se vier com ou sem máscara.
• O template Word deve conter os marcadores {{nome}}, {{cpf}}, {{curso}}.
• As certidões são convertidas para PDF e o arquivo .docx intermediário é removido.

------------------------------
5. Requisitos
------------------------------

• Python 3.x
• Pacotes:
    pip install pandas mysql-connector-python python-docx docx2pdf

• Word (apenas no Windows) para que o docx2pdf funcione

------------------------------
6. Autor
------------------------------
Este sistema foi desenvolvido por IGOR JOSÉ BATISTA DE BARROS, automatizando a emissão de documentos oficiais de forma confiável e segura.
