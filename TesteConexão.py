import mysql.connector
import datetime
from conn import config

# === CONFIGURAÇÃO DO BANCO === 
    # importado do arquivo conn.py

# Consulta SQL para buscar notas finais
query = """
SELECT
  u.id AS id_aluno,
  CONCAT(u.firstname, ' ', u.lastname) AS nome_aluno,
  c.fullname AS nome_curso,
  g.finalgrade AS nota_final,
  FROM_UNIXTIME(g.timemodified) AS data_avaliacao
FROM mdl_user u
JOIN mdl_user_enrolments ue ON ue.userid = u.id
JOIN mdl_enrol e ON e.id = ue.enrolid
JOIN mdl_course c ON c.id = e.courseid
JOIN mdl_grade_items gi ON gi.courseid = c.id AND gi.itemtype = 'course'
JOIN mdl_grade_grades g ON g.itemid = gi.id AND g.userid = u.id
WHERE g.finalgrade IS NOT NULL
ORDER BY u.lastname
LIMIT 20;
"""

try:
    # Conexão
    conn = mysql.connector.connect(**config)
    cursor = conn.cursor()

    # Execução da consulta
    cursor.execute(query)
    resultados = cursor.fetchall()

    # Exibição dos dados
    print("\n--- Histórico Escolar (Amostra) ---")
    for linha in resultados:
        id_aluno, nome_aluno, nome_curso, nota, data = linha
        print(f"Aluno: {nome_aluno}\nCurso: {nome_curso}\nNota Final: {nota:.2f}\nData: {data}\n{'-'*40}")

except mysql.connector.Error as err:
    print(f"Erro: {err}")

finally:
    try:
        if conn.is_connected():
            cursor.close()
            conn.close()
    except:
        pass
