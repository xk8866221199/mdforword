"""
Markdown to Word 文档转换应用
Flask Web 服务入口
"""
import os
import uuid
import tempfile
from flask import Flask, render_template, request, send_file, jsonify
from converter.docx_builder import convert_markdown_to_docx
from urllib.parse import quote

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# 临时文件目录
TEMP_DIR = os.path.join(tempfile.gettempdir(), 'mdforword')
os.makedirs(TEMP_DIR, exist_ok=True)


@app.route('/')
def index():
    """渲染主页面"""
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    """
    转换 Markdown 为 Word 文档
    接收 JSON: { "markdown": "..." , "filename": "..." }
    返回 JSON: { "download_id": "...", "filename": "..." }
    """
    try:
        data = request.get_json()
        if not data or 'markdown' not in data:
            return jsonify({'error': '请提供 Markdown 文本'}), 400

        markdown_text = data['markdown']
        filename = data.get('filename', '文档') or '文档'

        if not markdown_text.strip():
            return jsonify({'error': 'Markdown 文本不能为空'}), 400

        # 转换
        docx_buffer = convert_markdown_to_docx(markdown_text)

        # 清理文件名
        safe_filename = "".join(
            c for c in filename
            if c.isalnum() or c in (' ', '-', '_', '.', '（', '）')
            or '\u4e00' <= c <= '\u9fff'
        ).strip() or '文档'

        if not safe_filename.endswith('.docx'):
            safe_filename += '.docx'

        # 保存到临时文件
        download_id = str(uuid.uuid4())
        temp_path = os.path.join(TEMP_DIR, f'{download_id}.docx')
        with open(temp_path, 'wb') as f:
            f.write(docx_buffer.read())

        return jsonify({
            'download_id': download_id,
            'filename': safe_filename
        })

    except Exception as e:
        return jsonify({'error': f'转换失败: {str(e)}'}), 500


@app.route('/download/<download_id>')
def download(download_id):
    """
    通过 GET 请求下载已转换的文件
    浏览器原生处理下载，不受 JS 安全限制
    """
    # 安全检查 — 只允许 UUID 格式
    try:
        uuid.UUID(download_id)
    except ValueError:
        return jsonify({'error': '无效的下载链接'}), 400

    filename = request.args.get('name', '文档.docx')
    temp_path = os.path.join(TEMP_DIR, f'{download_id}.docx')

    if not os.path.exists(temp_path):
        return jsonify({'error': '文件不存在或已过期'}), 404

    response = send_file(
        temp_path,
        mimetype='application/vnd.openxmlformats-officedocument'
                 '.wordprocessingml.document',
        as_attachment=True,
        download_name=filename,
    )

    # 下载后删除临时文件（延迟删除，让 send_file 完成）
    @response.call_on_close
    def cleanup():
        try:
            os.remove(temp_path)
        except OSError:
            pass

    return response


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
