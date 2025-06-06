import tkinter as tk
from tkinter import filedialog
import tkinter.ttk as ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg,NavigationToolbar2Tk)
import os
import functools
import sqlite3

class JanelaPrincipal(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Análise de dados")
        self.geometry("400x300")
        self.config(bg="lightblue")

        self.db_path = "dados.db"
        self.bancoDados()

        self.cores = {
            "fundo": "#1E1E2E",
            "frame": "#2A2A3C",
            "centro": "#FFF",
            "texto": "#000",
            "destaque": "#00C6FF",
            "borda": "#44475A",
            "erro": "#FF5555"
        }

        self.frameTopo = tk.Frame(self)
        self.frameTopo.pack(side="top", pady=0, fill="x")

        self.frameMeio = tk.Frame(self)
        self.frameMeio.pack(expand=True, fill="both")


        self.framesEstatisticas = tk.Frame(self, bg=self.cores["frame"])

        self.framesOpcoes = tk.Frame(self, bg=self.cores["frame"], bd=2, relief="ridge")
        self.framesOpcoes.pack(side="bottom", fill="x", pady=2, ipady=2)

        self.tela_inicial()

    def bancoDados(self):
        if not os.path.exists(self.db_path):
            try:
                conn = sqlite3.connect(self.db_path)
                cur = conn.cursor()
                cur.execute('''
                    CREATE TABLE dados(
                        nome  TEXT PRIMARY KEY,
                        coluna TEXT,
                        valores FLOAT
                    )
                ''')
                conn.commit()
                conn.close()
                print("Banco e tabela criados com sucesso!")
            except Exception as e:
                print("Erro ao criar banco:", e)

    def temaClaro(self):
        self.configure(bg=self.cores["fundo"])
        self.frameTopo.configure(bg=self.cores["frame"])
        self.frameMeio.configure(bg=self.cores["centro"])
        self.framesOpcoes.configure(bg=self.cores["frame"])

        if hasattr(self, "resultado"):
            self.resultado.configure(bg=self.cores["centro"], fg=self.cores["texto"])
        if hasattr(self, "titulo"):
            self.titulo.configure(bg=self.cores["frame"], fg=self.cores["destaque"])
        if hasattr(self, "botao"):
            self.botao.configure(bg=self.cores["destaque"], fg="white")

        for widget in self.framesOpcoes.winfo_children():
            if isinstance(widget, tk.Button):
                widget.configure(bg=self.cores["destaque"], fg="white")

        for widget in self.frameMeio.winfo_children():
            if isinstance(widget, (tk.Label, tk.Button)):
                widget.configure(bg=self.cores["centro"], fg=self.cores["texto"])
                
    def tela_inicial(self):
        
        self.reset()

        self.titulo = tk.Label(self.frameTopo, text="Leitor de arquivos Excel", font=("Arial", 14))
        self.titulo.pack(pady=10)

        self.botao = tk.Button(self.frameMeio, text="Selecionar arquivo Excel", command=self.abrir_excel)
        self.botao.pack(pady=10)

        self.temaClaro()

        self.resultado = tk.Label(self.frameMeio, text="", justify="left", wraplength=500,bg=self.cores["centro"])
        self.resultado.pack(pady=20, ipady=20, ipadx=20)

        botoes = [
            ("Banco de dados", self.listarNomesBanco),
            ("Backups", lambda: print("Em desenvolvimento")),
            ("Novo Arquivo",  self.tela_inicial)
        ]

        for texto, funcao in botoes:
            btn = tk.Button(self.frameTopo, text=texto, command=funcao)
            btn.pack(side="left", padx=5, pady=5)

    def tabelaDados(self,tabela):
        self.reset()

        tabelaVizu = ttk.Treeview(self.frameMeio)
        tabelaVizu.pack(expand=True, fill="both")
        tabelaVizu["columns"] = list(tabela.columns)
        tabelaVizu["show"] = "headings"

        for col in tabela.columns:
            tabelaVizu.heading(col, text=col)
            larguraMaxima = max([len(str(valor)) for valor in tabela[col]] + [len(col)])
            pixels = larguraMaxima * 10
            tabelaVizu.column(col, width=pixels, anchor="center")

        for i, row in tabela.iterrows():
            if i >= 100:
                break
            tabelaVizu.insert("", "end", values=list(row))

    def abrir_excel(self):
        self.caminho = filedialog.askopenfilename(
            title="Selecione um arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if self.caminho:
            try:
                df = pd.read_excel(self.caminho)
                self.info =df.describe().rename(index={
                    "count": "Contagem",
                    "mean": "Média",
                    "std": "Desvio Padrão",
                    "min": "Mínimo",
                    "25%": "1º Quartil (25%)",
                    "50%": "Mediana (50%)",
                    "75%": "3º Quartil (75%)",
                    "max": "Máximo"
                }).round(2).to_string()
                self.opcoes(df)
                
            except Exception as e:
                self.resultado.config(text=f"Erro ao ler arquivo:\n{e}")

    def opcoes(self, tabela):
        self.reset()

        self.analiseBasica(tabela)

        botoes = [
            ("Análise básica", lambda: self.analiseBasica(tabela)),
            ("Tabela", lambda: self.tabelaDados(tabela)),
            ("Gerar Gráficos", lambda: self.selecionarTipo(tabela)),
            ("Estatísticas", lambda: self.estatisticas(tabela)),
            ("Filtros", lambda: self.selecionarFiltros(tabela))
        ]

        for index, (texto, funcao) in enumerate(botoes):
            linha = index // 3
            coluna = index % 3

            btn = tk.Button(self.framesOpcoes, text=texto, command=funcao)
            btn.grid(row=linha, column=coluna, padx=5, pady=5)

    def analiseBasica(self,tabela):
        self.reset()
        self.resultado.config(text=f'Análise básica:\n{self.info}')
    
    def estatisticas(self,tabela):
        self.reset()
        
        estatisticas_texto = ""
        for col in tabela.columns:
            if pd.api.types.is_numeric_dtype(tabela[col]):
                moda = str(tabela[col].mode().tolist()).replace("[", "").replace("]", "")
                unicos = tabela[col].nunique()
                estatisticas_texto += f"Coluna: {col}\nValores únicos: {unicos}\nModa: {moda}\n\n"
        correlacoes = tabela.corr(numeric_only=True)
        estatisticas_texto += f"Correlação entre colunas numéricas:\n{correlacoes.to_string()}"
        self.resultado.config(text=estatisticas_texto)
        
        
    def listarNomesBanco(self):
        try:
            
            con = sqlite3.connect(self.db_path)
            cur = con.cursor()
            cur.execute("SELECT nome FROM dados")
            nomes = cur.fetchall()
            self.reset()
            self.resultado.config(text=nomes)
            con.commit()
            con.close()
            
            

        except Exception as e:
            print(f'Erro: {e}')
    def vizualizarDados(self, nome):
        pass
    def salvar(self, dict):
        
        self.reset()
        self.resultado.config(text=dict)

        tk.Label(self.frameMeio, text="Digite o nome do conjunto de dados:").pack()
        nomeEntrada = tk.Entry(self.frameMeio)
        nomeEntrada.pack()
        
        def confirmar_salvamento():
            nome = nomeEntrada.get()
            try:
                conn = sqlite3.connect(self.db_path)
                cur = conn.cursor()
                for coluna, valores in dict.items():
                    for valor in valores:
                        cur.execute(
                        "INSERT OR REPLACE INTO dados (nome, coluna, valores) VALUES (?, ?, ?)",
                        (nome, coluna, str(valor))
                    )

                conn.commit()
                conn.close()
                self.resultado.config(text=f"Conjunto '{nome}' salvo com sucesso.")
            except Exception as e:
                self.resultado.config(text=f"Erro ao salvar: {e}")

        botaoSalvar = tk.Button(self.frameMeio, text="Salvar", command=confirmar_salvamento)
        botaoSalvar.pack(pady=10)

        
    def selecionarFiltros(self,tabela):
        self.reset()
        
        legenda = tk.Label(self.frameMeio, text="Escolha a coluna que irá filtrar os dados:")
        legenda.pack(pady=2)

        colunas = list(tabela.columns)
        self.colunaEscolhida = tk.StringVar(value=colunas[0])

        menu = tk.OptionMenu(self.frameMeio, self.colunaEscolhida, *colunas)
        menu.pack(pady=10, ipady=10)

        filtroBotao = tk.Button(self.frameMeio, text="Filtrar", command=lambda:self.analiseFiltro(tabela))
        
        filtroBotao.pack(pady=5)
        
        
    def selecionarTipo(self,tabela):
         self.reset()
         
         self.legendaSelecionadas={}
         self.colunasSelecionadas = {}  
         tabelaNumerica = [col for col in tabela.columns if pd.api.types.is_numeric_dtype(tabela[col])]
         legendas = [col for col in tabela.columns if pd.api.types.is_string_dtype(tabela[col])]

         tk.Label(self.frameMeio, text="Selecione as colunas numéricas para o gráfico").pack(pady=10)

         for coluna in tabelaNumerica:
              var = tk.BooleanVar()
              chk = tk.Checkbutton(self.frameMeio, text=coluna, variable=var)
              chk.pack(anchor='w')
              self.colunasSelecionadas[coluna] = var
         tk.Label(self.frameMeio, text="Selecione os textos para a legenda").pack(pady=10)

         for legenda in legendas:
              leg= tk.BooleanVar()
              legchk = tk.Checkbutton(self.frameMeio, text=legenda, variable=leg)
              legchk.pack(anchor='w')
              self.legendaSelecionadas[legenda] = leg
              
         tiposGraficos=[
         ("Barra", lambda: self.gerarGrafico("bar", tabela)),
         ("Histograma", lambda: self.gerarGrafico("hist",tabela)),
         ("Linha", lambda: self.gerarGrafico("line",tabela)),
         ("Dispersão", lambda: self.gerarGrafico("scatter",tabela))
         ]
         
          
         for index, (texto, funcao) in enumerate(tiposGraficos):
          

            btn = tk.Button(self.frameMeio, text=texto, command=funcao)
            btn.pack(side="left")
        
    def analiseFiltro(self,tabela):
        
        escolha = self.colunaEscolhida.get()
        if not pd.api.types.is_numeric_dtype(tabela[escolha]):
            self.filtrarTexto(tabela)
        else:
             media = tabela[escolha].mean()
             self.filtrarNumero(tabela)
            
    def filtrarNumero(self,tabela):
        self.reset()
    
        coluna = self.colunaEscolhida.get()
    
        tk.Label(self.frameMeio, text=f'Filtrar valores na coluna: {coluna}').pack(pady=10)

        tk.Label(self.frameMeio, text="Escolha o operador:").pack()
        self.operador = tk.StringVar(value="==")
        opcoes = ["==", ">", "<", ">=", "<=", "!="]
        menu = tk.OptionMenu(self.frameMeio, self.operador, *opcoes)
        menu.pack(pady=5)

        tk.Label(self.frameMeio, text="Digite o valor numérico:").pack()
        
        self.valorFiltro = tk.Entry(self.frameMeio)
        self.valorFiltro.pack(pady=5)
        
        botao= tk.Button(self.frameMeio, text="Aplicar", command= lambda: self.aplicarFiltro(tabela)).pack(pady=10)
        
    def filtrarTexto(self,tabela):
        self.reset()

        coluna = self.colunaEscolhida.get()
        titulo = tk.Label(self.frameMeio, text=f'Buscando valores na coluna: {coluna}')
        titulo.pack(pady=10)

        filtrarContem = tk.Label(self.frameMeio, text="Filtrar textos que contenham: ")
        filtrarContem.pack(pady=10)

        self.Contem = tk.Entry(self.frameMeio)
        self.Contem.pack(pady=10)

        botaoContem = tk.Button(self.frameMeio, text=" envio", command=lambda: self.contemTexto(tabela))
        botaoContem.pack(pady=2)

        filtrarInicio = tk.Label(self.frameMeio, text="Filtrar textos que comecem com: ")
        filtrarInicio.pack(pady=10)

        self.Inicio = tk.Entry(self.frameMeio)
        self.Inicio.pack(pady=10)

        botaoInicio = tk.Button(self.frameMeio, text="Envio", command=lambda: self.inicioTexto(tabela))
        botaoInicio.pack(pady=10)

        filtrarFim = tk.Label(self.frameMeio, text="Filtrar textos que terminem com: ")
        filtrarFim.pack(pady=10)

        self.Fim = tk.Entry(self.frameMeio)
        self.Fim.pack(pady=10)

        botaoFim = tk.Button(self.frameMeio, text="Envio", command=lambda: self.fimTexto(tabela))
        botaoFim.pack(pady=10)
        
    def contemTexto(self,tabela):
        
        col = self.colunaEscolhida.get()
        filtro = self.Contem.get()
        self.reset()
        
            
        try:
            resultado = tabela[tabela[col].str.contains(filtro, case=False, na=False)]

            
            dicionario_resultado={}
      
                
            for coluna in resultado.columns:
                dicionario_resultado[coluna] = resultado[coluna].tolist()
           
            self.salvar(dicionario_resultado)
           
        except Exception as e:
            self.resultado.config(text="Não existem, valores")

    def inicioTexto(self,tabela):
        
        col = self.colunaEscolhida.get()
        filtro = self.Inicio.get()
        
        self.reset()

        try:
            resultado = tabela[tabela[col].str.startswith(filtro)]
            self.resultado.config(text=resultado)
        except Exception as e:
            self.resultado.config(text=e)

    def fimTexto(self,tabela):
        
        col = self.colunaEscolhida.get()
        filtro = self.Fim.get()
        self.reset()

        try:
            resultado = tabela[tabela[col].str.endswith(filtro)]
            self.resultado.config(text=resultado)
            
        except Exception as e:
            self.resultado.config(text="Não existem valores")
            
    def aplicarFiltro(self,tabela):
        
        valor = self.valorFiltro.get()
        operador = self.operador.get()
        col = self.colunaEscolhida.get()
        self.reset()

        try:
            valor = float(valor)
            expressao = f"tabela[tabela['{col}'] {operador} {valor}]"
            tabelaFiltrada = eval(expressao, {"self": self, "__builtins__": {}})

            container = tk.Frame(self.frameMeio)
            container.pack(expand=True, fill="both")


            vsb = tk.Scrollbar(container, orient="vertical")
            hsb = tk.Scrollbar(container, orient="horizontal")

            tabelaVizu= ttk.Treeview(
            container,
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set
        )

            vsb.config(command=tabelaVizu.yview)
            hsb.config(command=tabelaVizu.xview)

            vsb.pack(side="right", fill="y")
            hsb.pack(side="bottom", fill="x")
            tabelaVizu.pack(side="left", expand=True, fill="both")

            tabelaVizu["columns"] = list(tabelaFiltrada.columns)
            tabelaVizu["show"] = "headings"

            for col in tabelaFiltrada.columns:
                tabelaVizu.heading(col, text=col)
                larguraMaxima = max([len(str(valor)) for valor in tabelaFiltrada[col]] + [len(col)])
                pixels = larguraMaxima * 10
                tabelaVizu.column(col, width=pixels, anchor="center")

            for i, row in tabelaFiltrada.iterrows():
                tabelaVizu.insert("", "end", values=list(row))

        except Exception as e:
            self.resultado.config(text="Erro ao aplicar filtro")
            
    def gerarGrafico(self, tipo, tabela):
        
        colunas_numericas = [col for col, var in self.colunasSelecionadas.items() if var.get()]
        colunas_legenda = [col for col, var in self.legendaSelecionadas.items() if var.get()]

        self.reset()

        if not colunas_numericas:
            self.resultado.config(text="Selecione ao menos uma coluna numérica para o gráfico.")
            return

        fig = Figure(figsize=(6, 4), dpi=100)
        ax = fig.add_subplot(111)

        try:
            if tipo == "bar":
                if colunas_legenda:
                    legenda = colunas_legenda[0]
                    for col in colunas_numericas:
                        agrupado = tabela.groupby(legenda)[col].mean()
                        ax.bar(agrupado.index, agrupado.values, label=col)
                    ax.set_xlabel(legenda)
                else:
                    for col in colunas_numericas:
                        ax.bar(tabela.index, tabela[col], label=col)
                ax.set_title("Gráfico de Barras")
                ax.legend()

            elif tipo == "line":
                for col in colunas_numericas:
                    ax.plot(tabela.index, tabela[col], label=col)
                ax.set_title("Gráfico de Linhas")
                ax.legend()

            elif tipo == "hist":
                for col in colunas_numericas:
                    ax.hist(tabela[col], bins=20, alpha=0.5, label=col)
                ax.set_title("Histograma")
                ax.legend()

            elif tipo == "scatter":
                if len(colunas_numericas) >= 2:
                    x = tabela[colunas_numericas[0]]
                    y = tabela[colunas_numericas[1]]
                    ax.scatter(x, y)
                    ax.set_xlabel(colunas_numericas[0])
                    ax.set_ylabel(colunas_numericas[1])
                    ax.set_title("Gráfico de Dispersão")
                else:
                    self.resultado.config(text="Selecione pelo menos duas colunas numéricas para gráfico de dispersão.")
                    return

            canvas = FigureCanvasTkAgg(fig, master=self.frameMeio)
            canvas.draw()
            canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

            toolbar = NavigationToolbar2Tk(canvas, self.frameMeio)
            toolbar.update()
            canvas.get_tk_widget().pack()

        except Exception as e:
            self.resultado.config(text=f"Erro ao gerar gráfico: {e}")

    def reset(self):
        
        if hasattr(self, "resultado") and self.resultado.winfo_exists():
            self.resultado.config(text="")
        for widget in self.frameMeio.winfo_children():
            if widget.winfo_exists() and widget != self.resultado:
                widget.destroy()
                
app = JanelaPrincipal()
app.mainloop()

