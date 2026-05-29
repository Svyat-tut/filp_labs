import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from letter_generator import LetterGenerator, Attachment
import os

class AttachmentDialog:

    def __init__(self, parent, attachment_num=1):
        self.result = None
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Приложение {attachment_num}")
        self.dialog.geometry("600x500")
        self.dialog.resizable(False, False)
        

        ttk.Label(self.dialog, text="Заголовок приложения:").pack(pady=5, padx=10, anchor=tk.W)
        self.title_entry = ttk.Entry(self.dialog, width=70)
        self.title_entry.pack(padx=10, pady=5)
        self.title_entry.insert(0, f"Содержание приложения {attachment_num}")
        

        ttk.Label(self.dialog, text="Текст приложения:").pack(pady=5, padx=10, anchor=tk.W)
        self.content_text = tk.Text(self.dialog, width=70, height=15, wrap=tk.WORD)
        self.content_text.pack(padx=10, pady=5)
        self.content_text.insert("1.0", "Введите текст приложения...")
        

        ttk.Label(self.dialog, text="Количество листов:").pack(pady=5, padx=10, anchor=tk.W)
        page_frame = ttk.Frame(self.dialog)
        page_frame.pack(padx=10, pady=5, anchor=tk.W)
        
        ttk.Label(page_frame, text="Лист(ов):").pack(side=tk.LEFT, padx=5)
        self.page_count = ttk.Spinbox(page_frame, from_=1, to=100, width=10)
        self.page_count.set(1)
        self.page_count.pack(side=tk.LEFT, padx=5)
        

        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(pady=20)
        
        ok_btn = ttk.Button(button_frame, text="ОК", command=self.ok_clicked)
        ok_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = ttk.Button(button_frame, text="Отмена", command=self.dialog.destroy)
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        self.dialog.grab_set()
    
    def ok_clicked(self):
        title = self.title_entry.get()
        content = self.content_text.get("1.0", tk.END)
        page_count = int(self.page_count.get())
        
        if not title.strip():
            messagebox.showwarning("Ошибка", "Заголовок не может быть пустым!")
            return
        
        self.result = Attachment(title, content, page_count)
        self.dialog.destroy()


class LetterGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор писем в DOCX с приложениями")
        self.root.geometry("700x700")
        self.root.resizable(False, False)
        
        self.letter_generator = LetterGenerator()
        self.attachments = []
        

        style = ttk.Style()
        style.theme_use('clam')
        

        title_label = ttk.Label(root, text="Генератор писем DOCX с приложениями", font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        

        form_frame = ttk.Frame(root, padding="10")
        form_frame.pack(fill=tk.BOTH, expand=True)
        

        ttk.Label(form_frame, text="ФИО адресата:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.recipient_name = ttk.Entry(form_frame, width=45)
        self.recipient_name.grid(row=0, column=1, pady=5, padx=10)
        self.recipient_name.insert(0, "Иван Иванович Иванов")
        

        ttk.Label(form_frame, text="Должность адресата:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.recipient_post = ttk.Entry(form_frame, width=45)
        self.recipient_post.grid(row=1, column=1, pady=5, padx=10)
        self.recipient_post.insert(0, "Директор компании")
        

        ttk.Label(form_frame, text="Организация:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.recipient_org = ttk.Entry(form_frame, width=45)
        self.recipient_org.grid(row=2, column=1, pady=5, padx=10)
        self.recipient_org.insert(0, "ООО Организация")
        

        ttk.Label(form_frame, text="Тема письма:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.letter_subject = ttk.Entry(form_frame, width=45)
        self.letter_subject.grid(row=3, column=1, pady=5, padx=10)
        self.letter_subject.insert(0, "Предложение о сотрудничестве")


        ttk.Label(form_frame, text="Текст письма:").grid(row=4, column=0, sticky=tk.NW, pady=5)
        self.letter_body = tk.Text(form_frame, width=45, height=6, wrap=tk.WORD)
        self.letter_body.grid(row=4, column=1, pady=5, padx=10)
        self.letter_body.insert("1.0", "Уважаемый коллега,\n\nПредлагаем вам ...")
        

        ttk.Label(form_frame, text="Приложения:").grid(row=5, column=0, sticky=tk.NW, pady=5)
        
        attachment_frame = ttk.Frame(form_frame)
        attachment_frame.grid(row=5, column=1, pady=5, padx=10, sticky=tk.W)
        
        add_attachment_btn = ttk.Button(attachment_frame, text="+ Добавить приложение", 
                                       command=self.add_attachment)
        add_attachment_btn.pack(side=tk.LEFT, padx=5)
        
        self.attachments_label = ttk.Label(attachment_frame, text="(нет приложений)", foreground="gray")
        self.attachments_label.pack(side=tk.LEFT, padx=10)
        

        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        generate_btn = ttk.Button(button_frame, text="Генерировать письмо", 
                                 command=self.generate_letter)
        generate_btn.pack(side=tk.LEFT, padx=5)
        
        template_btn = ttk.Button(button_frame, text="Выбрать шаблон", 
                                 command=self.select_template)
        template_btn.pack(side=tk.LEFT, padx=5)
        
        clear_attachments_btn = ttk.Button(button_frame, text="Очистить приложения", 
                                          command=self.clear_attachments)
        clear_attachments_btn.pack(side=tk.LEFT, padx=5)
        

        self.status_label = ttk.Label(form_frame, text="✓ Готово к работе", foreground="green")
        self.status_label.grid(row=7, column=0, columnspan=2, pady=10)
    
    def add_attachment(self):

        attachment_num = len(self.attachments) + 1
        dialog = AttachmentDialog(self.root, attachment_num)
        self.root.wait_window(dialog.dialog)
        
        if dialog.result:
            self.attachments.append(dialog.result)
            self.update_attachments_label()
            self.status_label.config(text=f"✓ Добавлено приложение ({len(self.attachments)} всего)", 
                                    foreground="blue")
    
    def clear_attachments(self):

        if self.attachments:
            self.attachments.clear()
            self.update_attachments_label()
            self.status_label.config(text="✓ Приложения очищены", foreground="orange")
    
    def update_attachments_label(self):

        if self.attachments:
            count = len(self.attachments)
            self.attachments_label.config(text=f"({count} приложение(й))", foreground="green")
        else:
            self.attachments_label.config(text="(нет приложений)", foreground="gray")
    
    def select_template(self):

        filename = filedialog.askopenfilename(
            title="Выберите шаблон DOCX",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
        )
        if filename:
            self.letter_generator.template_path = filename
            self.status_label.config(text=f"✓ Шаблон: {os.path.basename(filename)}", 
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
            

            output_path = self.letter_generator.generate(replacements, self.attachments)
            
            self.status_label.config(text=f"✓ Письмо создано: {os.path.basename(output_path)}", 
                                    foreground="green")
            messagebox.showinfo("Успех", f"Письмо создано успешно!\n\n{output_path}")
            
        except Exception as e:
            self.status_label.config(text="✗ Ошибка при создании письма", 
                                    foreground="red")
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = LetterGeneratorApp(root)
    root.mainloop()
