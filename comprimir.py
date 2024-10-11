import customtkinter as ctk
from tkinter import filedialog
from PIL import Image
import imageio  # Biblioteca alternativa para manipulação de vídeos
from docx import Document
from io import BytesIO
import threading
import time  # Importa a biblioteca de tempo

def compress_image_gui():
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.png")])
    if file_path:
        try:
            progress_bar.set(0)  # Inicializa a barra de progresso
            progress_bar.start()  # Inicia a barra de progresso

            img = Image.open(file_path)
            if img.mode == 'RGBA':
                img = img.convert('RGB')

            output_file = filedialog.asksaveasfilename(defaultextension=".jpg")
            if output_file:
                img.save(output_file, optimize=True, quality=50)
                show_message("Imagem comprimida com sucesso!", success=True)
        except Exception as e:
            show_message(f"Erro ao comprimir imagem: {e}")
        finally:
            progress_bar.stop()  # Para a barra de progresso

def compress_video_gui():
    file_path = filedialog.askopenfilename(filetypes=[("Video files", "*.mp4")])
    if file_path:
        output_path = filedialog.asksaveasfilename(defaultextension=".mp4")
        if output_path:
            # Inicia um thread separado para a compressão do vídeo
            threading.Thread(target=compress_video, args=(file_path, output_path)).start()

def compress_video(file_path, output_path):
    try:
        progress_bar.set(0)  # Inicializa a barra de progresso
        overall_progress_bar.set(0)  # Inicializa a barra de progresso geral
        progress_bar.start()  # Inicia a barra de progresso

        reader = imageio.get_reader(file_path)
        writer = imageio.get_writer(output_path, codec='libx264', bitrate='500k')

        total_frames = reader.count_frames()
        start_time = time.time()  # Tempo de início
        last_minutes = -1  # Inicializa a variável para controle dos minutos
        for i, frame in enumerate(reader):
            writer.append_data(frame)

            # Atualiza a barra de progresso a cada frame
            progress_bar.set((i + 1) / total_frames)

            # Atualiza a barra de progresso geral
            overall_progress_bar.set((i + 1) / total_frames)  # Atualiza a barra geral

            # Cálculo do tempo restante
            elapsed_time = time.time() - start_time  # Tempo gasto até agora
            estimated_total_time = (elapsed_time / (i + 1)) * total_frames  # Tempo estimado total
            remaining_time = estimated_total_time - elapsed_time  # Tempo restante

            # Atualiza a mensagem na textbox somente se mudar os minutos
            remaining_minutes, _ = divmod(remaining_time, 60)
            if remaining_minutes != last_minutes:  # Verifica se os minutos mudaram
                last_minutes = remaining_minutes  # Atualiza o último valor de minutos
                show_remaining_time(remaining_minutes)

        writer.close()
        reader.close()
        show_message("Vídeo comprimido com sucesso!", success=True)
    except Exception as e:
        show_message(f"Erro ao comprimir vídeo: {e}")
    finally:
        progress_bar.stop()  # Para a barra de progresso
        overall_progress_bar.stop()  # Para a barra de progresso geral

def compress_word_gui():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if file_path:
        try:
            progress_bar.set(0)  # Inicializa a barra de progresso
            progress_bar.start()  # Inicia a barra de progresso

            doc = Document(file_path)
            output_file = filedialog.asksaveasfilename(defaultextension=".docx")
            if output_file:
                total_images = sum(1 for rel in doc.part.rels.values() if "image" in rel.target_ref)
                if total_images == 0:
                    progress_bar.set(1.0)  # Se não houver imagens, completa a barra de progresso
                else:
                    for i, rel in enumerate(doc.part.rels.values()):
                        if "image" in rel.target_ref:
                            img_stream = BytesIO(rel.target_part.blob)
                            img = Image.open(img_stream)

                            output_stream = BytesIO()
                            img.save(output_stream, format='JPEG', quality=50)
                            output_stream.seek(0)

                            rel.target_part._blob = output_stream.getvalue()
                            progress_bar.set((i + 1) / total_images)  # Atualiza o progresso

                doc.save(output_file)
                show_message("Arquivo Word comprimido com sucesso!", success=True)
        except Exception as e:
            show_message(f"Erro ao comprimir arquivo Word: {e}")
        finally:
            progress_bar.stop()  # Para a barra de progresso

def show_remaining_time(minutes):
    # Limpa a linha anterior
    error_textbox.configure(state="normal")
    error_textbox.delete("1.0", "end")  # Limpa a textbox
    message = f"Tempo restante: {int(minutes)} min"
    error_textbox.insert("end", message)  # Insere a nova mensagem
    error_textbox.see("end")
    error_textbox.configure(state="disabled")  # Desabilita a textbox para evitar edições

def show_message(message, success=False):
    if success:
        message = f"✅ {message}"
    else:
        message = f"❌ {message}"
    error_textbox.configure(state="normal")
    error_textbox.insert("end", message + "\n")
    error_textbox.see("end")
    error_textbox.configure(state="disabled")  # Desabilita a textbox para evitar edições

# Inicializa o customtkinter
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title('Compressor de Mídia')

# Título
title_label = ctk.CTkLabel(root, text="Compressor de Mídia", font=("Arial", 20))
title_label.pack(pady=10)

# Barra de progresso para frames
ctk.CTkLabel(root, text="Progresso da Compressão do Vídeo", font=("Arial", 12)).pack(pady=(10, 0))
progress_bar = ctk.CTkProgressBar(root, width=300)
progress_bar.pack(pady=10)
progress_bar.set(0)

# Barra de progresso geral
ctk.CTkLabel(root, text="Progresso Geral", font=("Arial", 12)).pack(pady=(10, 0))
overall_progress_bar = ctk.CTkProgressBar(root, width=300)
overall_progress_bar.pack(pady=10)
overall_progress_bar.set(0)

# Botões
ctk.CTkButton(root, text="Comprimir Imagem (jpg e png)", command=compress_image_gui, font=("Arial", 13, "bold")).pack(pady=10, padx=10)
ctk.CTkButton(root, text="Comprimir Vídeo (Mp4)", command=compress_video_gui, font=("Arial", 13, "bold")).pack(pady=10)
ctk.CTkButton(root, text="Comprimir Word (docx)", command=compress_word_gui, font=("Arial", 13, "bold")).pack(pady=10)

# Textbox para mostrar mensagens
error_textbox = ctk.CTkTextbox(root, width=450, height=150, corner_radius=5)
error_textbox.pack(pady=10)
error_textbox.configure(state="normal")

# Inicia o loop principal da aplicação
root.mainloop()
