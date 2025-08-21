# Multicotas-Python
Multipool Olímpia

Sistema organizador de cotistas de multipropriedade, focado na visualização e gestão das datas de deixadas por clientes/proprietários de multipropriedades em nossa imobiliária para locação.
Oferecemos aqui um serviço que funciona da seguinte forma: nossa cidade é turística e conta com inúmeros resorts de multipropriedade, onde os proprietários que não estão satisfeitos com os serviços prestados pelos resorts, ou simplesmente nos encontram através de tráfego pago ou orgânico, os clientes/proprietários possuem datas disponíveis que normalmente são estipuladas por um cronograma de utilização pelo próprio empreendimento, nós realizamos todo o cadastro dessas e fazemos a locação para terceiros, onde cobramos uma taxa de até 30% se a locação é efetivada. 

Autor: Gabriel Recco Silva
Plataforma alvo: Windows (com OneDrive instalado para sincronização)

Principais recursos

Tecnologia utilizada Python

SQLite

Interface simples (PyQt5) com:

Abas: Datas Futuras, Datas Passadas, Histórico (logs) e Contabilidade (gráficos).

Barra de busca global (nome, data, nº da cota, apto, torre, etc.).

Ordenação rápida: Data mais próxima e Ordem alfabética.

CRUD completo (Adicionar/Editar/Excluir) com confirmação, leitura de .xlsx e exportação para Excel.

Bloqueio seguro de edição (lock-file + teste real de escrita no SQLite) → evita conflitos de multiusuário.

Backup automático do banco ao iniciar e logs detalhados por dia.

Alertas de próximos 7 dias e estatísticas em diálogo dedicado.

Dashboards (Matplotlib) com filtros de período e empreendimento.

Compatível com .py e .exe (PyInstaller) sem quebrar caminhos de arquivos.


Estrutura de pastas

.venv/   # Ambiente virtual (opcional, recomendado)

backups/         # Backups automáticos do SQLite (gerados pelo sistema)

build/           # Artefatos temporários do PyInstaller

dados/           # Banco de dados SQLite (multipool.db) e db.lock

dist/            # Executável gerado (versão .exe)

exportacoes/     # Planilhas .xlsx exportadas

logs/            # Arquivos de log (um por dia)

logo.png         # Ícone da janela

config.txt       # (opcional) Config (ver abaixo)

multipool_olimpia.py

O sistema cria as pastas necessárias na primeira execução.

*Configuração (opcional)*
O arquivo config.txt pode definir o caminho do banco:
DB_PATH=C:\Users\SeuUsuario\OneDrive\Multipool\dados\multipool.db

Sem config.txt, o sistema usa dados/multipool.db automaticamente.
O sistema funciona tanto em .py quanto .exe sem precisar editar caminhos.
Arquivos reconhecidos automaticamente:
logo.png → ícone da janela (se existir).
config.txt → configuração (se existir).

Requisitos (Usuários)

Windows
Acesso à Internet (para sincronização OneDrive)
OneDrive instalado/logado (se quiser sincronizar o banco e pastas)


Modo Somente Leitura (multiusuário)

Ao abrir, o sistema verifica se há outro usuário editando (lock real no SQLite):
Se bloqueado, você entra em Modo Leitura (consulta sem alterações).
Se livre, o sistema cria um arquivo dados/db.lock para sinalizar edição.
Ao fechar, o lock é removido automaticamente (em caso de falha, ele é validado na próxima execução).

Atalhos úteis
Ctrl+N: Adicionar
F2: Editar
Delete: Excluir
Ctrl+S: Exportar para Excel
Ctrl+O: Importar Excel (.xlsx)
Ctrl+D: Ordenar por data
Ctrl+A: Ordenar por cotista
Ctrl+F: Focar pesquisa
F5: Recarregar dados
Ctrl+1 / … / Ctrl+4: Alternar abas
Ctrl+Q: Sair

Importação e Exportação

Exportação gera um .xlsx em exportacoes/ com todos os campos principais e internos.
Importação aceita .xlsx (openpyxl) com cabeçalhos na ordem:

Cotista, Contato, Empreendimento, Entrada, Saída, Dormitório, Valor,
Disponível, Fonte, Nº da Cota, Nº Apartamento, Torre, Letra de Prioridade

Datas aceitas: dd/MM/yyyy ou yyyy-MM-dd (normalização automática).
Valores: qualquer texto numérico (o sistema lida com R$ e vírgula/ponto).
Deduplicação básica por Cotista + Entrada + Empreendimento.


Contabilidade (dashboards)

Filtros: Data inicial, Data final, Empreendimento.
Gráficos:
Registros por mês (12 últimos)
Disponibilidade (Sim/Não)
Fontes (Cliente / Lead Internet / Terceiros)
Valores por mês (R$)
Resumo acima dos gráficos: Qtd. de registros e Valor Total.

Logs & Backups
Backups: criados automaticamente em backups/ ao iniciar.
Logs: gravados em logs/log_YYYY-MM-DD.txt a cada ação (inserir/atualizar/excluir/exportar/importar).

Licença

Uso interno da Multicotas Olímpia

Entre em contato para termos específicos de distribuição.

