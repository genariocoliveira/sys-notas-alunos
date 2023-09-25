import tkinter as tk
from tkinter import ttk
import pandas as pd

class PrincipalRAD:
    def __init__(self, win):
        self.win = win
        self.win.title('Bem Vindo ao RAD')
        self.win.geometry("820x600+10+10")

        self.dadosColunas = ("Aluno", "Nota1", "Nota2", "Média", "Situação")

        # Componentes
        self.lblNome = tk.Label(win, text='Nome do Aluno:')
        self.lblNota1 = tk.Label(win, text='Nota 1')
        self.lblNota2 = tk.Label(win, text='Nota 2')
        self.lblMedia = tk.Label(win, text='Média')

        self.txtNome = tk.Entry(win, bd=3)
        self.txtNota1 = tk.Entry(win)
        self.txtNota2 = tk.Entry(win)

        self.btnCalcular = tk.Button(win, text='Calcular Média', command=self.fCalcularMedia)
        self.btnSalvar = tk.Button(win, text='Salvar Dados', command=self.fSalvarDados)

        self.treeMedias = ttk.Treeview(win, columns=self.dadosColunas, selectmode='browse')
        self.verscrlbar = ttk.Scrollbar(win, orient="vertical", command=self.treeMedias.yview)
        self.treeMedias.configure(yscrollcommand=self.verscrlbar.set)

        self.treeMedias.heading("Aluno", text="Aluno")
        self.treeMedias.heading("Nota1", text="Nota 1")
        self.treeMedias.heading("Nota2", text="Nota 2")
        self.treeMedias.heading("Média", text="Média")
        self.treeMedias.heading("Situação", text="Situação")

        self.treeMedias.column("Aluno", minwidth=0, width=100)
        self.treeMedias.column("Nota1", minwidth=0, width=100)
        self.treeMedias.column("Nota2", minwidth=0, width=100)
        self.treeMedias.column("Média", minwidth=0, width=100)
        self.treeMedias.column("Situação", minwidth=0, width=100)

        # Posicionamento dos componentes
        self.lblNome.place(x=100, y=50)
        self.txtNome.place(x=200, y=50)

        self.lblNota1.place(x=100, y=100)
        self.txtNota1.place(x=200, y=100)

        self.lblNota2.place(x=100, y=150)
        self.txtNota2.place(x=200, y=150)

        self.btnCalcular.place(x=100, y=200)
        self.btnSalvar.place(x=220, y=200)

        self.treeMedias.place(x=100, y=300)
        self.verscrlbar.place(x=805, y=300, height=225)

        self.iid = 0

        # Carregar dados iniciais ao iniciar o programa
        self.carregarDadosIniciais()

    def carregarDadosIniciais(self):
        try:
            fsave = 'planilhaAlunos.xlsx'
            dados = pd.read_csv(fsave)
            print('Dados carregados com sucesso de', fsave)

            for i, row in dados.iterrows():
                nome = row['Aluno']
                nota1 = str(row['Nota1'])
                nota2 = str(row['Nota2'])
                media = str(row['Média'])
                situacao = row['Situação']

                self.treeMedias.insert('', 'end',
                                       values=(nome, nota1, nota2, media, situacao))

                self.iid += 1

        except FileNotFoundError:
            print('Arquivo de dados não encontrado. Comece a adicionar dados manualmente.')

    def fSalvarDados(self):
        try:
            fsave = 'planilhaAlunos.xlsx'

            dados = []
            for line in self.treeMedias.get_children():
                lstDados = []
                for value in self.treeMedias.item(line)['values']:
                    lstDados.append(value)
                dados.append(lstDados)

            df = pd.DataFrame(data=dados, columns=self.dadosColunas)
            df.to_csv(fsave, index=False)
            print('Dados salvos com sucesso em', fsave)

        except Exception as e:
            print('Erro ao salvar dados:', str(e))

    def fCalcularMedia(self):
        try:
            nome = self.txtNome.get()
            nota1 = float(self.txtNota1.get())
            nota2 = float(self.txtNota2.get())
            media, situacao = self.fVerificarSituacao(nota1, nota2)

            self.treeMedias.insert('', 'end',
                                   values=(nome,
                                           str(nota1),
                                           str(nota2),
                                           str(media),
                                           situacao))

        except ValueError:
            print('Entre com valores válidos')
        finally:
            self.txtNome.delete(0, 'end')
            self.txtNota1.delete(0, 'end')
            self.txtNota2.delete(0, 'end')

    def fVerificarSituacao(self, nota1, nota2):
        media = (nota1 + nota2) / 2
        if media >= 7.0:
            situacao = 'Aprovado'
        elif media >= 5.0:
            situacao = 'Em Recuperação'
        else:
            situacao = 'Reprovado'

        return media, situacao

if __name__ == "__main__":
    janela = tk.Tk()
    principal = PrincipalRAD(janela)
    janela.mainloop()
