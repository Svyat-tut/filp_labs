import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from letter_generator import LetterGenerator
import os

class LetterGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор писем в DOCX")
        self.root.geometry("600x500")
        self.root.resizable(False, False)
        
        self.letter_generator = LetterGenerator()
        

        style = ttk.Style()
        style.theme_use('clam')
        

        title_label = ttk.Label(root, text="Генератор писем DOCX", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        

        form_frame = ttk.Frame(root, padding="20")
        form_frame.pack(fill=tk.BOTH, expand=True)
        

        ttk.Label(form_frame, text="ФИО адресата:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.recipient_name = ttk.Entry(form_frame, width=40)
        self.recipient_name.grid(row=0, column=1, pady=5, padx=10)
        self.recipient_name.insert(0, "Иван Иванович Иванов")
        

        ttk.Label(form_frame, text="Должность адресата:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.recipient_post = ttk.Entry(form_frame, width=40)
        self.recipient_post.grid(row=1, column=1, pady=5, padx=10)
        self.recipient_post.insert(0, "Директор компании")
        

        ttk.Label(form_frame, text="Организация:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.recipient_org = ttk.Entry(form_frame, width=40)
        self.recipient_org.grid(row=2, column=1, pady=5, padx=10)
        self.recipient_org.insert(0, "ООО Организация")
        

        ttk.Label(form_frame, text="Тема письма:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.letter_subject = ttk.Entry(form_frame, width=40)
        self.letter_subject.grid(row=3, column=1, pady=5, padx=10)
        self.letter_subject.insert(0, "Предложение о сотрудничестве")
        

        ttk.Label(form_frame, text="Текст письма:").grid(row=4, column=0, sticky=tk.NW, pady=5)
        self.letter_body = tk.Text(form_frame, width=40, height=6, wrap=tk.WORD)
        self.letter_body.grid(row=4, column=1, pady=5, padx=10)
        self.letter_body.insert("1.0", "Уважаемый коллега,\n\nПредлагаем вам ...")
        

        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=20)
        
        generate_btn = ttk.Button(button_frame, text="Генерировать письмо", 
                                 command=self.generate_letter)
        generate_btn.pack(side=tk.LEFT, padx=5)
        
        template_btn = ttk.Button(button_frame, text="Выбрать шаблон", 
                                 command=self.select_template)
        template_btn.pack(side=tk.LEFT, padx=5)


        self.status_label = ttk.Label(form_frame, text="Готово к работе", 
                                      foreground="green")
        self.status_label.grid(row=6, column=0, columnspan=2, pady=10)
    
    def select_template(self):
        filename = filedialog.askopenfilename(
            title="Выберите шаблон DOCX",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
        )
        if filename:
            self.letter_generator.template_path = filename
            self.status_label.config(text=f"Шаблон: {os.path.basename(filename)}", 
                                    foreground="blue")
    
    def generate_letter(self):
        try:
            replacements = {
                "{RECIPIENT_NAME}": self.recipient_name.get(),
                "{RECIPIENT_POST}": self.recipient_post.get(),
                "{RECIPIENT_ORG}": self.recipient_org.get(),
                "{LETTER_SUBJECT}": self.letter_subject.get(),
                "{LETTER_BODY}": self.letter_body.get("1.0", tk.END)
            }
            

            output_path = self.letter_generator.generate(replacements)
            
            self.status_label.config(text=f"✓ Письмо создано: {os.path.basename(output_path)}", 
                                    foreground="green")
            messagebox.showinfo("Успех", f"Письмо создано успешно!\n\n{output_path}")
            
        except Exception as e:
            self.status_label.config(text="Ошибка при создании письма", 
                                    foreground="red")
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = LetterGeneratorApp(root)
    root.mainloop()
