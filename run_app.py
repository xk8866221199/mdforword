"""
MD → Word 桌面应用
使用 pywebview 封装 Flask Web 应用为 macOS 原生窗口
"""
import threading
import sys
import os
import shutil


def _get_resource_path():
    """获取资源文件路径（兼容打包和开发模式）"""
    if getattr(sys, 'frozen', False):
        # py2app 打包后，资源在 .app/Contents/Resources/
        return os.path.dirname(sys.executable).replace(
            '/MacOS', '/Resources'
        )
    else:
        # 开发模式，资源在脚本同目录
        return os.path.dirname(os.path.abspath(__file__))


class Api:
    """暴露给 JavaScript 的 Python API"""

    def __init__(self, window_ref, temp_dir):
        self._window = window_ref
        self._temp_dir = temp_dir

    def save_file(self, download_id, filename):
        """
        弹出 macOS 原生保存对话框，让用户选择保存位置
        由 JavaScript 调用: window.pywebview.api.save_file(id, name)
        """
        import webview

        # 获取窗口引用
        window = webview.windows[0] if webview.windows else None
        if not window:
            return {'success': False, 'error': '窗口未找到'}

        # 临时文件路径
        temp_path = os.path.join(self._temp_dir, f'{download_id}.docx')
        if not os.path.exists(temp_path):
            return {'success': False, 'error': '文件不存在或已过期'}

        # 弹出原生保存文件对话框
        try:
            save_path = window.create_file_dialog(
                webview.SAVE_DIALOG,
                directory=os.path.expanduser('~/Documents'),
                save_filename=filename,
                file_types=('Word 文档 (*.docx)',),
            )
        except Exception as e:
            return {'success': False, 'error': f'对话框错误: {str(e)}'}

        if save_path:
            # save_path 可能是字符串或元组
            target = save_path if isinstance(save_path, str) else save_path[0]
            if not target.endswith('.docx'):
                target += '.docx'
            try:
                shutil.copy2(temp_path, target)
                # 清理临时文件
                os.remove(temp_path)
                return {'success': True, 'path': target}
            except Exception as e:
                return {'success': False, 'error': f'保存失败: {str(e)}'}
        else:
            return {'success': False, 'error': '用户取消保存'}


def main():
    resource_dir = _get_resource_path()

    # 设置 Flask 模板和静态文件路径
    template_dir = os.path.join(resource_dir, 'templates')
    static_dir = os.path.join(resource_dir, 'static')

    # 切换工作目录到资源目录
    os.chdir(resource_dir)

    # 导入并配置 Flask app
    from app import app
    app.template_folder = template_dir
    app.static_folder = static_dir

    import webview
    import tempfile

    # 临时文件目录
    temp_dir = os.path.join(tempfile.gettempdir(), 'mdforword')
    os.makedirs(temp_dir, exist_ok=True)

    # 创建暴露给 JS 的 API
    api = Api(window_ref=None, temp_dir=temp_dir)

    def start_flask():
        """在后台线程中启动 Flask 服务器"""
        app.run(
            host='127.0.0.1',
            port=5001,
            debug=False,
            use_reloader=False,
        )

    # 启动 Flask 后台线程
    flask_thread = threading.Thread(target=start_flask, daemon=True)
    flask_thread.start()

    # 创建原生 macOS 窗口
    window = webview.create_window(
        title='MD → Word',
        url='http://127.0.0.1:5001',
        width=1200,
        height=800,
        min_size=(800, 600),
        confirm_close=False,
        js_api=api,
    )

    # 启动 WebView（macOS 使用原生 WebKit）
    webview.start()

    # 窗口关闭后退出
    sys.exit(0)


if __name__ == '__main__':
    main()
