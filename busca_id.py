import mysql.connector
import pandas as pd
import re
import os
from conn import config


# Parâmetros de busca
busca_aluno = 'jorge luiz vanderlei de araujo'
id_curso = ''  # Altere para a turma desejada
role_id_aluno = 5  # ID padrão de aluno no Moodle

try:
    conn = mysql.connector.connect(**config)
    cursor = conn.cursor(dictionary=True)

    if busca_aluno.strip():
        cursor.execute("""
            SELECT 
                u.id AS id_aluno,
                CONCAT(u.firstname, ' ', u.lastname) AS nome_completo,
                u.username,
                c.fullname AS curso,
                c.id AS id_curso
            FROM mdl_user u
            JOIN mdl_user_enrolments ue ON ue.userid = u.id
            JOIN mdl_enrol e ON e.id = ue.enrolid
            JOIN mdl_course c ON c.id = e.courseid
            JOIN mdl_role_assignments ra ON ra.userid = u.id
            WHERE CONCAT(u.firstname, ' ', u.lastname) LIKE %s
            AND ra.roleid = %s
        """, (f'%{busca_aluno.strip()}%', role_id_aluno))
        alunos = cursor.fetchall()

        if alunos:
            df = pd.DataFrame(alunos)
            os.makedirs("Certidoes_Emitidas", exist_ok=True)
            caminho_arquivo = os.path.join("Certidoes_Emitidas", "aluno_busca_nome.xlsx")
            df.to_excel(caminho_arquivo, index=False)
            print(f"✅ {len(df)} aluno(s) exportado(s) para '{caminho_arquivo}'")
        else:
            print("Nenhum aluno encontrado com esse nome.")
    else:
        # Primeiro, descobrir o contextid do curso
        cursor.execute("""
            SELECT ctx.id AS contextid, c.fullname AS curso
            FROM mdl_context ctx
            JOIN mdl_course c ON c.id = ctx.instanceid
            WHERE ctx.contextlevel = 50 AND c.id = %s
        """, (id_curso,))
        context = cursor.fetchone()

        if not context:
            print("❌ Curso não encontrado.")
        else:
            contextid = context['contextid']
            nome_curso = context['curso']
            nome_curso_limpo = re.sub(r'[\\/*?:"<>|]', "_", nome_curso)  # limpa nome do arquivo

            # Agora buscamos somente os usuários com role de aluno (roleid = 5)
            query = """
                SELECT 
                    u.id AS id_aluno,
                    CONCAT(u.firstname, ' ', u.lastname) AS nome_completo,
                    u.username,
                    c.fullname AS curso,
                    c.id AS id_curso
                FROM mdl_user u
                JOIN mdl_user_enrolments ue ON ue.userid = u.id
                JOIN mdl_enrol e ON e.id = ue.enrolid
                JOIN mdl_course c ON c.id = e.courseid
                JOIN mdl_role_assignments ra ON ra.userid = u.id
                WHERE c.id = %s
                  AND ra.contextid = %s
                  AND ra.roleid = %s
            """

            cursor.execute(query, (id_curso, contextid, role_id_aluno))
            alunos = cursor.fetchall()

            if alunos:
                df = pd.DataFrame(alunos)
                nome_arquivo = f"alunos_{nome_curso_limpo}.xlsx"
                # Garante que a pasta de saída existe
                os.makedirs("Certidoes_Emitidas", exist_ok=True)
                # Define o caminho completo para salvar na pasta
                caminho_arquivo = os.path.join("Certidoes_Emitidas", f"alunos_{nome_curso_limpo}.xlsx")
                df.to_excel(caminho_arquivo, index=False)
                print(f"✅ {len(df)} aluno(s) exportados para '{caminho_arquivo}'")
            else:
                print("⚠️ Nenhum aluno encontrado com papel de estudante neste curso.")

except mysql.connector.Error as err:
    print(f"Erro: {err}")

finally:
    if conn.is_connected():
        cursor.close()
        conn.close()
