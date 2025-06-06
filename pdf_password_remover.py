import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import tempfile
import traceback
from pathlib import Path

# 必要なライブラリの事前チェック
missing_libraries = []

try:
    import fitz  # PyMuPDF
except ImportError:
    missing_libraries.append("PyMuPDF")

try:
    from PIL import Image
except ImportError:
    missing_libraries.append("Pillow")

try:
    import pypdf
except ImportError:
    missing_libraries.append("pypdf")

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
except ImportError:
    missing_libraries.append("reportlab")

# EXE実行時のパス調整
if getattr(sys, 'frozen', False):
    # EXE実行時
    application_path = sys._MEIPASS
else:
    # 通常のPython実行時
    application_path = os.path.dirname(os.path.abspath(__file__))

class PDFPasswordRemover:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Password Remover v1.0")
        self.root.geometry("750x550")
        
        # 変数の初期化
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.password = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.extraction_method = tk.StringVar(value="copy")
        
        self.processing = False
        self.japanese_font = 'Helvetica'
        
        self.create_widgets()
        self.setup_japanese_font()
        
        # 起動時のライブラリチェック
        self.check_libraries()
    
    def check_libraries(self):
        """起動時に必要なライブラリをチェック"""
        if missing_libraries:
            error_msg = "以下のライブラリが不足しています:\n" + "\n".join(missing_libraries)
            error_msg += "\n\nこのアプリを正常に動作させるには、これらのライブラリが必要です。"
            messagebox.showerror("ライブラリエラー", error_msg)
            return False
        return True
    
    def setup_japanese_font(self):
        """日本語フォントの設定"""
        try:
            import platform
            system = platform.system()
            
            if system == "Windows":
                font_paths = [
                    "C:/Windows/Fonts/msgothic.ttc",
                    "C:/Windows/Fonts/meiryo.ttc",
                    "C:/Windows/Fonts/NotoSansCJK-Regular.ttc",
                    "C:/Windows/Fonts/yugothm.ttc"
                ]
            elif system == "Darwin":  # macOS
                font_paths = [
                    "/System/Library/Fonts/Hiragino Sans GB.ttc",
                    "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
                    "/System/Library/Fonts/NotoSansCJK.ttc",
                    "/Library/Fonts/Arial Unicode MS.ttf"
                ]
            else:  # Linux
                font_paths = [
                    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
                    "/usr/share/fonts/truetype/takao-gothic/TakaoGothic.ttf",
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
                ]
            
            self.japanese_font = 'Helvetica'  # デフォルト
            
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont('Japanese', font_path))
                        self.japanese_font = 'Japanese'
                        self.log_message(f"日本語フォントを設定: {os.path.basename(font_path)}")
                        break
                    except Exception as font_error:
                        continue
                        
            if self.japanese_font == 'Helvetica':
                self.log_message("日本語フォントが見つかりません。Helveticaを使用します。")
                
        except Exception as e:
            self.japanese_font = 'Helvetica'
            self.log_message(f"フォント設定エラー: {e}")
        
    def create_widgets(self):
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 入力ファイル選択
        ttk.Label(main_frame, text="パスワード保護されたPDFファイル:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_file, width=60).grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="参照", command=self.browse_input_file).grid(row=1, column=1, padx=(5, 0), pady=5)
        
        # パスワード入力
        ttk.Label(main_frame, text="パスワード:").grid(row=2, column=0, sticky=tk.W, pady=5)
        password_frame = ttk.Frame(main_frame)
        password_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5, columnspan=2)
        
        self.password_entry = ttk.Entry(password_frame, textvariable=self.password, show="*", width=50)
        self.password_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.show_password = tk.BooleanVar()
        ttk.Checkbutton(password_frame, text="表示", variable=self.show_password, 
                       command=self.toggle_password_visibility).pack(side=tk.RIGHT, padx=(5, 0))
        
        # 抽出方法選択
        ttk.Label(main_frame, text="抽出方法:").grid(row=4, column=0, sticky=tk.W, pady=5)
        method_frame = ttk.Frame(main_frame)
        method_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=5, columnspan=2)
        
        methods = [
            ("copy", "ページコピー (推奨: 高速・高品質)", "最も確実で高速な方法"),
            ("image", "画像抽出 (確実・大容量)", "完全な見た目保持、確実だが重い"),
        ]
        
        for value, text, desc in methods:
            frame = ttk.Frame(method_frame)
            frame.pack(anchor=tk.W, fill=tk.X)
            
            rb = ttk.Radiobutton(frame, text=text, variable=self.extraction_method, value=value)
            rb.pack(side=tk.LEFT)
            
            ttk.Label(frame, text=f"  ({desc})", foreground="gray").pack(side=tk.LEFT)
        
        # 出力ファイル選択
        ttk.Label(main_frame, text="出力先PDFファイル:").grid(row=6, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_file, width=60).grid(row=7, column=0, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="参照", command=self.browse_output_file).grid(row=7, column=1, padx=(5, 0), pady=5)
        
        # 実行ボタン
        self.convert_button = ttk.Button(main_frame, text="PDFを変換", command=self.convert_pdf)
        self.convert_button.grid(row=8, column=0, pady=15)
        
        # プログレスバー
        ttk.Label(main_frame, text="進行状況:").grid(row=9, column=0, sticky=tk.W, pady=5)
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=10, column=0, sticky=(tk.W, tk.E), pady=5, columnspan=2)
        
        # ログエリア
        ttk.Label(main_frame, text="ログ:").grid(row=11, column=0, sticky=tk.W, pady=5)
        
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=12, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.log_text = tk.Text(log_frame, height=10, width=80, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # グリッドの重み設定
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(12, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 初期メッセージ
        self.log_message("PDF Password Remover v1.0 が起動しました。")
        self.log_message("1. パスワード保護されたPDFファイルを選択してください。")
    
    def toggle_password_visibility(self):
        """パスワードの表示/非表示を切り替え"""
        if self.show_password.get():
            self.password_entry.configure(show="")
        else:
            self.password_entry.configure(show="*")
        
    def browse_input_file(self):
        filename = filedialog.askopenfilename(
            title="パスワード保護されたPDFファイルを選択",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            # 出力ファイル名を自動設定
            if not self.output_file.get():
                base_name = os.path.splitext(filename)[0]
                self.output_file.set(f"{base_name}_unlocked.pdf")
            
    def browse_output_file(self):
        filename = filedialog.asksaveasfilename(
            title="出力PDFファイルを保存",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
            
    def log_message(self, message):
        """ログメッセージを表示"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def set_processing_state(self, processing):
        """処理中状態の設定"""
        self.processing = processing
        state = "disabled" if processing else "normal"
        self.convert_button.configure(state=state)
        
        if not processing:
            self.progress_var.set(0)
        
    def validate_inputs(self):
        """入力値の検証"""
        if not self.input_file.get():
            messagebox.showerror("エラー", "入力ファイルを選択してください")
            return False
            
        if not os.path.exists(self.input_file.get()):
            messagebox.showerror("エラー", "入力ファイルが存在しません")
            return False
            
        if not self.password.get():
            messagebox.showerror("エラー", "パスワードを入力してください")
            return False
            
        if not self.output_file.get():
            messagebox.showerror("エラー", "出力ファイルを選択してください")
            return False
                
        return True
        
    def convert_pdf(self):
        """PDFの変換を実行"""
        if self.processing:
            return
            
        if not self.validate_inputs():
            return
            
        self.set_processing_state(True)
        
        try:
            method = self.extraction_method.get()
            method_names = {
                "copy": "ページコピー方式",
                "image": "画像抽出方式",
            }
            
            self.log_message(f"開始: {method_names.get(method, method)}")
            self.log_message(f"入力: {os.path.basename(self.input_file.get())}")
            self.log_message(f"出力: {os.path.basename(self.output_file.get())}")
            self.log_message("-" * 50)
            
            # 変換方法に応じて処理
            if method == "copy":
                self.convert_via_copy()
            elif method == "image":
                self.convert_via_images()
                
            self.progress_var.set(100)
            self.log_message("-" * 50)
            self.log_message("✓ 変換が正常に完了しました！")
            
            # ファイルサイズの表示
            if os.path.exists(self.output_file.get()):
                size = os.path.getsize(self.output_file.get())
                size_mb = size / (1024 * 1024)
                self.log_message(f"出力ファイルサイズ: {size_mb:.2f} MB")
            
            messagebox.showinfo("完了", "PDFの変換が完了しました")
            
        except Exception as e:
            error_msg = str(e)
            self.log_message(f"✗ エラーが発生しました: {error_msg}")
            self.log_message("\n詳細なエラー情報:")
            self.log_message(traceback.format_exc())
            messagebox.showerror("エラー", f"変換中にエラーが発生しました:\n{error_msg}")
            
        finally:
            self.set_processing_state(False)
    
    def convert_via_copy(self):
        """ページコピー方式（最も推奨）"""
        self.log_message("ページをそのままコピーします...")
        
        try:
            with open(self.input_file.get(), 'rb') as input_file:
                reader = pypdf.PdfReader(input_file)
                
                self.log_message(f"ページ数: {len(reader.pages)}")
                
                if reader.is_encrypted:
                    self.log_message("パスワードを確認中...")
                    if not reader.decrypt(self.password.get()):
                        raise Exception("パスワードが正しくありません")
                    self.log_message("✓ パスワード認証成功")
                
                writer = pypdf.PdfWriter()
                
                for page_num in range(len(reader.pages)):
                    self.log_message(f"ページ {page_num + 1}/{len(reader.pages)} をコピー中...")
                    
                    page = reader.pages[page_num]
                    writer.add_page(page)
                    
                    progress = (page_num + 1) / len(reader.pages) * 95
                    self.progress_var.set(progress)
                
                self.log_message("ファイルを保存中...")
                with open(self.output_file.get(), 'wb') as output_file:
                    writer.write(output_file)
                    
        except Exception as e:
            if "password" in str(e).lower():
                raise Exception("パスワードが正しくありません")
            else:
                raise Exception(f"ページコピー処理エラー: {str(e)}")
    
    def convert_via_images(self):
        """画像抽出方式"""
        self.log_message("各ページを高解像度画像として抽出します...")
        
        doc = None
        try:
            doc = fitz.open(self.input_file.get())
            
            if doc.needs_pass:
                self.log_message("パスワードを確認中...")
                if not doc.authenticate(self.password.get()):
                    raise Exception("パスワードが正しくありません")
                self.log_message("✓ パスワード認証成功")
                
            self.log_message(f"ページ数: {len(doc)}")
            
            with tempfile.TemporaryDirectory() as temp_dir:
                image_files = []
                
                for page_num in range(len(doc)):
                    self.log_message(f"ページ {page_num + 1}/{len(doc)} を画像に変換中...")
                    
                    page = doc.load_page(page_num)
                    
                    # 高解像度マトリックス (300 DPI)
                    mat = fitz.Matrix(300/72, 300/72)
                    pix = page.get_pixmap(matrix=mat)
                    
                    image_path = os.path.join(temp_dir, f"page_{page_num + 1:03d}.png")
                    pix.save(image_path)
                    image_files.append(image_path)
                    
                    progress = (page_num + 1) / len(doc) * 80
                    self.progress_var.set(progress)
                
                self.log_message("画像からPDFを作成中...")
                self.create_pdf_from_images(image_files, self.output_file.get())
                
        except Exception as e:
            if "password" in str(e).lower():
                raise Exception("パスワードが正しくありません")
            else:
                raise Exception(f"画像抽出処理エラー: {str(e)}")
        finally:
            if doc:
                doc.close()
    
    def create_pdf_from_images(self, image_files, output_path):
        """画像ファイルのリストからPDFを作成"""
        if not image_files:
            raise Exception("変換する画像がありません")
            
        images = []
        try:
            # 全ての画像を開いてRGB変換
            for img_path in image_files:
                img = Image.open(img_path)
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                images.append(img)
            
            if not images:
                raise Exception("有効な画像がありません")
            
            # 最初の画像を使ってPDFを作成
            first_image = images[0]
            other_images = images[1:] if len(images) > 1 else []
            
            first_image.save(
                output_path,
                format='PDF',
                save_all=True,
                append_images=other_images,
                resolution=300.0,
                quality=95
            )
            
        except Exception as e:
            raise Exception(f"PDF作成エラー: {str(e)}")
        finally:
            # メモリ解放
            for img in images:
                try:
                    img.close()
                except:
                    pass

def main():
    """メイン関数"""
    try:
        root = tk.Tk()
        app = PDFPasswordRemover(root)
        
        # ウィンドウアイコンの設定（オプション）
        try:
            root.iconbitmap("icon.ico")  # アイコンファイルがある場合
        except:
            pass
            
        root.mainloop()
        
    except Exception as e:
        import tkinter as tk
        import tkinter.messagebox as messagebox
        try:
            root = tk.Tk()
            root.withdraw()  # メインウィンドウを隠す
            messagebox.showerror("起動エラー", f"アプリケーションの起動に失敗しました:\n{str(e)}")
        except:
            print(f"起動エラー: {str(e)}")

if __name__ == "__main__":
    main()
