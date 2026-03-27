pipeline {
    agent any

    environment {
        ROOT_DIR = '/fileserver05/SFTP/KRUK/DEMANDAS'
    }

    stages {

        stage('Crear venv') {
            steps {
                sh 'python3.11 --version'
                sh 'rm -rf venv'
                sh 'python3.11 -m venv venv'
            }
        }

        stage('1 - Naming CSV') {
            steps {
                sh "venv/bin/python3.11 naming_csv.py \"${ROOT_DIR}\""
            }
        }

        stage('2 - Stamp PDF') {
            steps {
                sh "venv/bin/python3.11 stamp_pdf.py \"${ROOT_DIR}\""
            }
        }

        stage('3 - Sign PDF') {
            steps {
                sh "venv/bin/python3.11 sign_pdf.py \"${ROOT_DIR}\""
            }
        }

        stage('4 - Convertir rutas a Windows') {
            steps {
                sh """
venv/bin/python3.11 - << 'PYEOF'
import openpyxl, os

def to_win(p):
    return p.replace("/fileserver05", "K:").replace("/", "\\\\")

root = "${ROOT_DIR}"
for fname in ["documentos.xlsx", "demandas.xlsx"]:
    path = os.path.join(root, fname)
    if not os.path.isfile(path):
        print(f"No encontrado: {fname}")
        continue
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    try:
        ruta_col = headers.index("ruta") + 1
    except ValueError:
        print(f"Columna 'ruta' no encontrada en {fname}")
        continue
    for row in ws.iter_rows(min_row=2):
        cell = row[ruta_col - 1]
        if cell.value:
            cell.value = to_win(str(cell.value))
    wb.save(path)
    print(f"Rutas convertidas: {fname}")
PYEOF
"""
            }
        }
    }

    post {
        success { echo 'Proceso completado correctamente.' }
        failure  { echo 'El proceso ha fallado. Revisa los logs de cada stage.' }
    }
}
