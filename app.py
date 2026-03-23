import os
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import io
from web_processor import process_excel_file

app = Flask(__name__)
# Chave secreta necessária para usar o flash (mensagens de erro na tela)
app.secret_key = os.environ.get('SECRET_KEY', 'minha_chave_super_secreta_alares')

# Extensões permitidas
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # Verifica se a requisição tem a parte do arquivo
    if 'file' not in request.files:
        flash('Nenhum arquivo enviado.')
        return redirect(request.url)
    
    file = request.files['file']
    
    # Se o usuário não selecionar nenhum arquivo
    if file.filename == '':
        flash('Nenhum arquivo selecionado.')
        return redirect(url_for('index'))
        
    if file and allowed_file(file.filename):
        try:
            # Lê o arquivo em memória para um BytesIO
            input_stream = io.BytesIO(file.read())
            
            # Repassa para o nosso processador
            output_stream = process_excel_file(input_stream)
            
            # Mantém o nome original, adicionando um sufixo
            original_name = file.filename
            name_part = original_name.rsplit('.', 1)[0]
            new_filename = f"{name_part}_Processado.xlsx"
            
            # Devolve o arquivo para download
            return send_file(
                output_stream,
                as_attachment=True,
                download_name=new_filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
        except ValueError as ve:
            # Erros de validação (ex: falta da aba 3)
            flash(str(ve))
            return redirect(url_for('index'))
        except Exception as e:
            # Erros gerais
            flash(f"Ocorreu um erro técnico: {str(e)}")
            return redirect(url_for('index'))
    else:
        flash('Tipo de arquivo não permitido. Envie apenas planilhas Excel (.xlsx ou .xls).')
        return redirect(url_for('index'))

if __name__ == '__main__':
    # Roda o servidor. Porta padrão 5000.
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
