import subprocess
import json
import docx
import re
import os
from datetime import datetime
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import pymupdf
import ast

APP_TITLE = "Validador de Currículos"
RESULTS_FILE = "results.json"
DOCUMENT_TYPE = "CURRÍCULO"
REQUIREMENTS = [
    "Coerência geral",
    "Introdução / Resumo profissional",
    "Formação Acadêmica",
    "Experiência Profissional"
]
JSON_REQUIREMENTS = [
    { "item": "Coerência geral", "status": "", "detalhes": "" },
    { "item": "Introdução / Resumo profissional", "status": "", "detalhes": "" },
    { "item": "Formação Acadêmica", "status": "", "detalhes": "" },
    { "item": "Experiência Profissional", "status": "", "detalhes": "" }
]
MODELS = ["llama3.1:8b", "deepseek-r1:8b", "gpt-oss:20b", "gemma3:12b", ""]

# --- Core Resume Processing Functions ------------------------------------------------

def extract_text_from_file(path):
    # Detecta tipo de arquivo
    ext = os.path.splitext(path)[1].lower()

    if ext == ".docx":
        return extract_text_from_docx(path)
    elif ext == ".pdf":
        return extract_text_from_pdf(path)

    else:
        return {
            "error": f"Formato não suportado: {ext}",
            "file": os.path.basename(path)
        }
     
def extract_text_from_docx(path):
    doc = docx.Document(path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

def extract_text_from_pdf(path):
    text = []
    with pymupdf.open(path) as pdf:
        for page in pdf:
            text.append(page.get_text())
    return "\n".join(text)

def build_prompt(text, requirements):
    req_text = "\n".join([f"- {r}" for r in requirements])
    return f"""
Analise o {DOCUMENT_TYPE} abaixo segundo os requisitos a seguir:

REQUISITOS:
{req_text}

FORMATO OBRIGATÓRIO DA RESPOSTA:
Retorne APENAS um JSON válido (RFC 8259).
Use exclusivamente aspas duplas.
NÃO use aspas simples.
NÃO adicione texto antes ou depois

{{
  "validacao":{JSON_REQUIREMENTS},
  "pontuacao_final": "0-100",
  "melhorias_recomendadas": ""
}}

{DOCUMENT_TYPE}:
{text}
"""

def ollama_chat(model, prompt):
    """
    Calls ollama locally.
    Returns raw text output.
    """
    try:
        result = subprocess.run(
            ["ollama", "run", model, "--think=false"],
            input=prompt.encode("utf-8"),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=300
        )
        if result.returncode != 0:
            # include stderr for debugging
            err = result.stderr.decode("utf-8", errors="ignore")
            out = result.stdout.decode("utf-8", errors="ignore")
            return out + "\n\n[OLLAMA STDERR]\n" + err
        return result.stdout.decode("utf-8", errors="ignore")
    except FileNotFoundError:
        return "[ERROR] 'ollama' executable not found. Ensure ollama is installed and in PATH."
    except subprocess.TimeoutExpired:
        return "[ERROR] ollama run timed out."
    except Exception as e:
        return f"[ERROR] ollama run failed: {e}"

def extract_json(text):
    """
    Extract JSON object from the model output. Accepts code block or raw JSON.
    """
    match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if match:
        return match.group(1)

    start = text.find("{")
    if start != -1:
        potential_json = text[start:].strip()
        # find matching closing brace for outermost object by scanning
        depth = 0
        for i, ch in enumerate(potential_json):
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    return potential_json[:i+1]
    return None

def save_result_entry(entry):
    """
    Append entry into RESULTS_FILE (create if missing).
    """
    if not os.path.exists(RESULTS_FILE):
        with open(RESULTS_FILE, "w", encoding="utf-8") as f:
            json.dump([entry], f, indent=2, ensure_ascii=False)
    else:
        try:
            with open(RESULTS_FILE, "r", encoding="utf-8") as f:
                content = json.load(f)
        except Exception:
            content = []
        content.append(entry)
        with open(RESULTS_FILE, "w", encoding="utf-8") as f:
            json.dump(content, f, indent=2, ensure_ascii=False)

def load_history():
    if not os.path.exists(RESULTS_FILE):
        return []
    try:
        with open(RESULTS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def clear_history_file():
    if os.path.exists(RESULTS_FILE):
        os.remove(RESULTS_FILE)
    selection = self.history_list.curselection()
    if not selection:
        return

    index = selection[0]
    entry = self.history[index]
    file_path = entry.get("file")

    if file_path and os.path.exists(file_path):
        os.startfile(file_path)  # Windows
    else:
        messagebox.showerror("Erro", "Arquivo não encontrado.")

def safe_json_loads(s):
    """
    Aceita JSON inválido (aspas simples, etc) e converte para dict usando ast.literal_eval.
    """
    try:
        return json.loads(s)
    except Exception:
        try:
            return ast.literal_eval(s)  # aceita aspas simples, sintaxe tipo Python
        except Exception as e:
            raise ValueError(f"JSON inválido: {e}")

def validate_resume_local(path, model="llama3.1:8b"):
    """
    Synchronous function that runs the whole pipeline and returns a dict result.
    """
    text = extract_text_from_file(path)
    prompt = build_prompt(text, REQUIREMENTS)
    response = ollama_chat(model, prompt)
    json_str = extract_json(response)

    if not json_str:
        data = {"error": "Nenhum JSON encontrado", "raw": response}
    else:
        try:
            data = safe_json_loads(json_str)
        except Exception as e:
            data = {
                "error": f"Falha ao converter JSON: {e}",
                "json_extraido": json_str,
                "raw": response
            }

    entry = {
        "file": os.path.basename(path),
        "path": os.path.abspath(path),
        "timestamp": datetime.now().isoformat(),
        "model": model,
        "result": data
    }
    save_result_entry(entry)
    return entry

# --- PDF Export Utility --------------------------------------------------------------

def export_entry_to_pdf(entry, outpath):
    """
    Try to export entry to PDF using reportlab. If reportlab not installed, raise ImportError.
    """
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    except Exception as e:
        raise ImportError("reportlab not available") from e

    doc = SimpleDocTemplate(outpath, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []

    header = f"Análise de Currículo - {entry.get('file','')}"
    story.append(Paragraph(header, styles['Title']))
    meta = f"Arquivo: {entry.get('path','') or entry.get('file','')}<br/>Data: {entry.get('timestamp','')}<br/>Modelo: {entry.get('model','')}"
    story.append(Paragraph(meta, styles['Normal']))
    story.append(Spacer(1, 12))

    # Add JSON prettified
    json_text = json.dumps(entry.get("result", {}), indent=2, ensure_ascii=False)
    for line in json_text.splitlines():
        # small paragraphs for each line to keep layout simple
        story.append(Paragraph(line.replace(" ", "&nbsp;"), styles['Code'] if 'Code' in styles else styles['Normal']))

    doc.build(story)

# --- Threaded Worker -----------------------------------------------------------------

def worker_analyze(path, model, result_queue):
    """
    Worker that runs validate_resume_local and puts entry into result_queue.
    """
    try:
        entry = validate_resume_local(path, model)
        result_queue.put(("ok", entry))
    except Exception as e:
        result_queue.put(("error", str(e)))

# --- GUI ------------------------------------------------------------------------------

class ResumeAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1100x650")
        self.queue = queue.Queue()

        # Theme state
        self.dark_mode = False
        self.setup_styles()

        # Top frame: controls
        top_frame = ttk.Frame(root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=8)

        # Model selector
        ttk.Label(top_frame, text="Modelo:").pack(side=tk.LEFT, padx=(0,6))
        self.model_var = tk.StringVar(value="llama3.1:8b")
        models = MODELS
        self.model_combo = ttk.Combobox(top_frame, values=models, textvariable=self.model_var, width=22)
        self.model_combo.pack(side=tk.LEFT, padx=(0,10))

        # Select file button
        self.btn_select = ttk.Button(top_frame, text="Selecionar Arquivo (.docx, .pdf)", command=self.select_file)
        self.btn_select.pack(side=tk.LEFT, padx=(0,10))

        # Delete selected
        self.delete_button = ttk.Button(top_frame, text="Excluir Selecionado", command=self.delete_selected)
        self.delete_button.pack(side=tk.LEFT, padx=(0,10))

        # Re-run selected
        self.btn_rerun = ttk.Button(top_frame, text="Rodar novamente", command=self.rerun_selected)
        self.btn_rerun.pack(side=tk.LEFT, padx=(0,10))

        # Export
        self.btn_export = ttk.Button(top_frame, text="Exportar selecionado (PDF/TXT)", command=self.export_selected)
        self.btn_export.pack(side=tk.LEFT, padx=(0,10))

        # Clear history
        self.btn_clear = ttk.Button(top_frame, text="Limpar histórico", command=self.clear_history_confirm)
        self.btn_clear.pack(side=tk.LEFT, padx=(0,10))

        # Theme toggle
        self.theme_btn = ttk.Button(top_frame, text="Modo Escuro", command=self.toggle_theme)
        self.theme_btn.pack(side=tk.RIGHT, padx=(0,6))

        # Progress bar
        self.progress = ttk.Progressbar(root, mode="indeterminate")
        self.progress.pack(fill=tk.X, padx=8, pady=(0,8))

        # Main frames: left history, right details
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        left = ttk.Frame(main_frame, width=280)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0,8))

        ttk.Label(left, text="Histórico (results.json):").pack(anchor=tk.W)
        self.history_list = tk.Listbox(left, width=40, activestyle='dotbox')
        self.history_list.pack(fill=tk.Y, expand=True)
        self.history_list.bind("<<ListboxSelect>>", self.on_history_select)
        self.history_list.bind("<Double-1>", self.open_selected_file)

        # Buttons under list
        hb = ttk.Frame(left)
        hb.pack(fill=tk.X, pady=(6,0))
        ttk.Button(hb, text="Atualizar", command=self.reload_history).pack(side=tk.LEFT, padx=(0,6))
        ttk.Button(hb, text="Abrir pasta", command=self.open_results_folder).pack(side=tk.LEFT)

        # Right side: result display
        right = ttk.Frame(main_frame)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        meta_frame = ttk.Frame(right)
        meta_frame.pack(fill=tk.X)
        self.meta_label = ttk.Label(meta_frame, text="Selecione um item do histórico ou faça uma nova análise", anchor=tk.W)
        self.meta_label.pack(fill=tk.X, padx=2, pady=2)

        self.result_text = scrolledtext.ScrolledText(right, font=("Consolas", 11), wrap=tk.WORD)
        self.result_text.tag_configure("title", font=("Consolas", 14, "bold"))
        self.result_text.tag_configure("bold", font=("Consolas", 11, "bold"))
        self.result_text.pack(fill=tk.BOTH, expand=True)

        # Last row: status
        self.status_var = tk.StringVar(value="Pronto")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Load history initially
        self.history = []
        self.reload_history()

        # Poll queue for worker results
        self.root.after(200, self.check_queue)

    # UI / Theme -------------------------------------------------------------------
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('default')
        # Styles
        self.bg_light = "#f0f0f0"
        self.fg_light = "#000000"
        self.bg_dark = "#2b2b2b"
        self.fg_dark = "#ffffff"
        self.update_theme()

    def update_theme(self):
        if self.dark_mode:
            bg = self.bg_dark
            fg = self.fg_dark
            self.theme_btn_text = "Modo Claro"
        else:
            bg = self.bg_light
            fg = self.fg_light
            self.theme_btn_text = "Modo Escuro"

        self.root.configure(bg=bg)
        try:
            for child in self.root.winfo_children():
                child.configure(bg=bg)
        except Exception:
            pass
        # update button text safely
        try:
            self.theme_btn.config(text=self.theme_btn_text)
        except Exception:
            pass

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.update_theme()

    # History management -----------------------------------------------------------
    def reload_history(self):
        self.history = load_history()
        self.history_list.delete(0, tk.END)
        for i, item in enumerate(self.history):
            ts = item.get("timestamp", "")
            fname = item.get("file", "unknown")
            model = item.get("model", "")
            label = f"{i+1}. {fname} — {ts.split('T')[0]} — {model}"
            self.history_list.insert(tk.END, label)
        self.status_var.set(f"Histórico carregado: {len(self.history)} itens")

    def open_results_folder(self):
        folder = os.getcwd()
        try:
            if os.name == "nt":
                os.startfile(folder)
            elif os.name == "posix":
                subprocess.Popen(["xdg-open", folder])
            else:
                messagebox.showinfo("Pasta", f"Pasta atual: {folder}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta: {e}")

    def clear_history_confirm(self):
        if messagebox.askyesno("Confirmar", "Deseja limpar completamente o histórico (results.json)?"):
            try:
                clear_history_file()
                self.reload_history()
                self.result_text.delete(1.0, tk.END)
                self.result_text.tag_configure("bold", font=("Arial", 10, "bold"))
                self.meta_label.config(text="Histórico limpo")
                self.status_var.set("Histórico limpo")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível limpar: {e}")

    def open_selected_file(self, event=None):
        selection = self.history_list.curselection()
        if not selection:
            return

        index = selection[0]
        entry = self.history[index]
        file_path = entry.get("file")

        if file_path and os.path.exists(file_path):
            os.startfile(file_path)  # Windows
        else:
            messagebox.showerror("Erro", "Arquivo não encontrado.")

    # Selection / display ---------------------------------------------------------
    def on_history_select(self, event=None):
        sel = self.history_list.curselection()
        if not sel:
            return
        idx = sel[0]
        entry = self.history[idx]
        self.display_entry(entry)

    def display_entry(self, entry):
        self.meta_label.config(
            text=f"{entry.get('file','')} — {entry.get('timestamp','')} — {entry.get('model','')}"
        )

        result = entry.get("result", {})
        self.result_text.delete(1.0, tk.END)

        if "error" in result:
            self.result_text.insert(tk.END, json.dumps(result, indent=2, ensure_ascii=False))
            return

        lines = []

        validacao = result.get("validacao", [])
        for item in validacao:
            titulo = item.get("item", "Item")
            status = item.get("status", "")
            detalhes = item.get("detalhes", "")

            lines.append((f"{titulo}\n", "title"))

            lines.append(("Status: ", "bold"))
            lines.append((status + "\n", None))

            lines.append(("Detalhes: ", "bold"))
            lines.append((detalhes + "\n\n", None))

        lines.append(("Pontuação Final:\n", "title"))
        lines.append((f"{result.get('pontuacao_final', '')}\n\n", None))

        melhorias = result.get("melhorias_recomendadas", "")
        if melhorias:
            lines.append(("Melhorias Recomendadas:\n", "title"))
            lines.append((f"{melhorias}\n", None))

        # Send to text box
        for text, tag in lines:
            if tag:
                self.result_text.insert(tk.END, text, tag)
            else:
                self.result_text.insert(tk.END, text)

        self.status_var.set(f"Exibindo item: {entry.get('file','')}")


    # File selection and analysis -----------------------------------------------
    def select_file(self):
        path = filedialog.askopenfilename(
            title="Selecione o arquivo (.docx, .pdf)",
            filetypes=[
                ("Documentos", "*.docx;*.pdf"),
                ("Word", "*.docx"),
                ("PDF", "*.pdf")
            ]
        )
        if not path:
            return
        self.run_analysis(path)

    def delete_selected(self):
        selection = self.history_list.curselection()
        if not selection:
            messagebox.showerror("Erro", "Nenhum item selecionado.")
            return

        index = selection[0]

        confirm = messagebox.askyesno("Confirmar", "Deseja excluir este resultado?")
        if not confirm:
            return

        # remove from memory
        self.history.pop(index)

        # save results.json
        with open(RESULTS_FILE, "w", encoding="utf-8") as f:
            json.dump(self.history, f, indent=2, ensure_ascii=False)

        # update listbox
        self.history_list.delete(index)

        # clear the display panel
        self.result_text.delete(1.0, tk.END)
        self.meta_label.config(text="")

    def rerun_selected(self):
        sel = self.history_list.curselection()
        if not sel:
            messagebox.showinfo("Re-run", "Nenhum item selecionado no histórico.")
            return
        idx = sel[0]
        entry = self.history[idx]
        path = entry.get("path") or entry.get("file")
        if not path or not os.path.exists(path):
            # allow user to locate file
            messagebox.showinfo("Arquivo não encontrado", "O arquivo original não foi encontrado. Selecione um arquivo para reexecutar.")
            self.select_file()
            return
        self.run_analysis(path)

    def run_analysis(self, path):
        model = self.model_var.get().strip() or "llama3.1:8b"
        # start progress and thread
        self.progress.start(10)
        self.status_var.set(f"Analisando {os.path.basename(path)} com {model} ...")
        self.btn_select.config(state=tk.DISABLED)
        self.delete_button.config(state=tk.DISABLED)
        self.btn_rerun.config(state=tk.DISABLED)
        self.btn_export.config(state=tk.DISABLED)
        self.btn_clear.config(state=tk.DISABLED)
        self.theme_btn.config(state=tk.DISABLED)

        t = threading.Thread(target=worker_analyze, args=(path, model, self.queue), daemon=True)
        t.start()

    def check_queue(self):
        try:
            item = self.queue.get_nowait()
        except queue.Empty:
            self.root.after(200, self.check_queue)
            return

        status, payload = item
        self.progress.stop()
        self.btn_select.config(state=tk.NORMAL)
        self.delete_button.config(state=tk.NORMAL)
        self.btn_rerun.config(state=tk.NORMAL)
        self.btn_export.config(state=tk.NORMAL)
        self.btn_clear.config(state=tk.NORMAL)
        self.theme_btn.config(state=tk.NORMAL)

        if status == "ok":
            entry = payload
            # reload history and auto-select last item
            self.reload_history()
            # select last item
            last_index = len(self.history) - 1
            if last_index >= 0:
                self.history_list.select_clear(0, tk.END)
                self.history_list.select_set(last_index)
                self.history_list.event_generate("<<ListboxSelect>>")
            messagebox.showinfo("Concluído", f"Análise concluída e salva: {entry.get('file')}")
            self.status_var.set("Análise concluída")
        else:
            err = payload
            messagebox.showerror("Erro na análise", f"Ocorreu um erro: {err}")
            self.status_var.set("Erro na análise")

        self.root.after(200, self.check_queue)

    # Exporting ------------------------------------------------------------------
    def export_selected(self):
        sel = self.history_list.curselection()
        if not sel:
            messagebox.showinfo("Exportar", "Nenhum item selecionado no histórico.")
            return
        entry = self.history[sel[0]]
        # ask where to save
        default_name = os.path.splitext(entry.get("file", "result"))[0] + "_analise.pdf"
        out_path = filedialog.asksaveasfilename(title="Salvar como...", defaultextension=".pdf", initialfile=default_name, filetypes=[("PDF", "*.pdf"),("Text", "*.txt")])
        if not out_path:
            return
        try:
            if out_path.lower().endswith(".pdf"):
                try:
                    export_entry_to_pdf(entry, out_path)
                    messagebox.showinfo("Exportado", f"Exportado para PDF: {out_path}")
                except ImportError:
                    # fallback: save as txt
                    txt_path = out_path[:-4] + ".txt"
                    with open(txt_path, "w", encoding="utf-8") as f:
                        f.write(json.dumps(entry, indent=2, ensure_ascii=False))
                    messagebox.showinfo("Fallback TXT", f"reportlab não encontrado. Resultado salvo como TXT: {txt_path}")
            else:
                with open(out_path, "w", encoding="utf-8") as f:
                    f.write(json.dumps(entry, indent=2, ensure_ascii=False))
                messagebox.showinfo("Exportado", f"Exportado como texto: {out_path}")
        except Exception as e:
            messagebox.showerror("Erro exportar", f"Erro ao exportar: {e}")

def main():
    root = tk.Tk()
    app = ResumeAnalyzerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
