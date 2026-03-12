⚡ AUTOMAÇÃO POWER BI + EXCEL
Google Drive → Power BI Desktop
Versão - Com clique automático e log visual HTML

O que essa automação faz?
	•	Monitora automaticamente qualquer alteração salva no seu arquivo Excel
	•	Detecta novas abas e colunas adicionadas na planilha
	•	Fecha e reabre o Power BI Desktop com os dados atualizados
	•	🆕 Clica em 'Atualizar' automaticamente — sem precisar tocar no mouse
	•	🆕 Gera um log visual em HTML com histórico de todas as atualizações
	•	Faz backup automático do Excel antes de cada atualização
	•	Pode ser agendada para rodar em horários fixos ou a cada X minutos
	•	Inicia sozinha toda vez que você ligar o computador

1. Arquivos do pacote
O pacote contém 5 arquivos. Coloque todos na mesma pasta:

Arquivo
Função
automacao_powerbi.py
Script principal da automação (não edite este arquivo)
config.json
Todas as configurações: caminhos, horários, ativar/desativar funções
instalar_dependencias.bat
Instala as bibliotecas Python com 1 clique
iniciar.bat
Inicia a automação manualmente
agendar_tarefa.ps1
Configura a automação para iniciar com o Windows

2. Pré-requisitos
✓
Python 3.8+
Baixe em python.org/downloads — marque 'Add Python to PATH' na instalação
✓
Google Drive Desktop
Baixe em google.com/drive/download — sincroniza o Excel localmente
✓
Power BI Desktop
Baixe pelo Microsoft Store ou powerbi.microsoft.com
✓
Arquivo .pbix existente
Seu relatório já deve existir e apontar para o Excel como fonte de dados

3. Instalação (5 minutos)
Passo 1 — Salvar os arquivos
Coloque todos os 5 arquivos em uma pasta fixa no seu computador. Exemplo:
C:\Automacao\PowerBI\

Passo 2 — Instalar dependências
Clique duas vezes no arquivo abaixo. Ele instala tudo automaticamente, incluindo o pyautogui para o clique automático:
instalar_dependencias.bat
Aguarde a mensagem: [OK] Todas as dependencias instaladas!

Passo 3 — Configurar o config.json
Abra o config.json com o Bloco de Notas e preencha os 3 campos obrigatórios:

Campo
O que preencher
excel_path
Caminho do Excel no Google Drive local.
Ex: C:\Users\joao\Google Drive\planilha.xlsx
powerbi_pbix_path
Caminho do seu arquivo .pbix.
Ex: C:\Users\joao\Documents\relatorio.pbix
powerbi_exe_path
Onde está instalado o Power BI Desktop.
Ex: C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe

4. Como executar
Opção A — Manual
Clique duas vezes no iniciar.bat. Uma janela de terminal ficará aberta monitorando o Excel. Quando você salvar o arquivo, a automação dispara sozinha.

Opção B — Automático com o Windows (recomendado)
Abra o PowerShell como Administrador e rode:
PowerShell -ExecutionPolicy Bypass -File "C:\Automacao\PowerBI\agendar_tarefa.ps1"
A automação vai iniciar sozinha toda vez que você ligar o computador, sem precisar fazer nada.

5. Clique automático em 'Atualizar' (novidade v2)
Após abrir o Power BI, o script aguarda o modelo carregar e envia o atalho de teclado Alt + H + R para acionar o botão Atualizar automaticamente.

Como funciona na prática
	•	O script abre o Power BI Desktop com o seu arquivo .pbix
	•	Aguarda o número de segundos configurado em 'aguardar_carregamento_pbi' (padrão: 20s)
	•	Foca a janela do Power BI e envia o atalho de teclado
	•	Aguarda o refresh concluir ('aguardar_refresh_segundos', padrão: 15s)
	•	Registra o resultado no log HTML

Configurações do clique automático no config.json
"auto_click": {
"ativo": true,
"metodo": "teclado",
"aguardar_refresh_segundos": 15
}

⚠️  Importante: se o atalho não funcionar
O atalho Alt+H+R funciona na maioria das instalações. Se não funcionar:
	•	Aumente o valor de 'aguardar_carregamento_pbi' — o PBI pode estar demorando mais para abrir
	•	Verifique se o Power BI está em português — o atalho pode variar por idioma
	•	Em último caso, mude 'ativo': false no auto_click — o PBI abrirá normalmente e você clica manualmente
O restante da automação (monitorar, detectar mudanças, backup, log) continuará funcionando normalmente.

6. Log visual em HTML (novidade v2)
A cada atualização o script gera o arquivo logs\historico.html dentro da pasta da automação. Abra esse arquivo no navegador para acompanhar tudo.

O que o log mostra
	•	Cards coloridos: verde = sucesso, amarelo = aviso (ex: nova coluna detectada), vermelho = erro
	•	Data e hora exata de cada atualização
	•	Nome do arquivo Excel monitorado
	•	Mudanças detectadas: novas abas, novas colunas, colunas removidas
	•	Contadores no topo: total de atualizações OK, avisos e erros
	•	A página recarrega automaticamente a cada 30 segundos
	•	Histórico de até 200 eventos é mantido entre sessões

Para abrir o log: navegue até a pasta da automação → entre na pasta logs → clique duas vezes em historico.html.

7. Onde encontrar o caminho do Google Drive
O Google Drive Desktop cria uma pasta local sincronizada. Para achar o caminho do seu Excel:

	•	Abra o Explorador de Arquivos
	•	No menu lateral, procure 'Google Drive' ou 'Meu Drive'
	•	Navegue até o seu Excel, clique com botão direito
	•	Clique em 'Copiar como caminho' e cole no config.json

Exemplos de caminho:
G:\Meu Drive\planilha.xlsx
C:\Users\joao\Google Drive\planilha.xlsx

8. Configurações avançadas no config.json
⏰ Horários fixos de atualização
Para atualizar às 8h, 12h e 18h, edite assim:
"horarios_fixos": ["08:00", "12:00", "18:00"]
Quando horarios_fixos estiver preenchido, o campo intervalo_minutos é ignorado.

⏱️  Ajustar tempo de espera (datasets grandes)
Se o seu dataset é grande e o Power BI demora para carregar, aumente esses valores:
"aguardar_carregamento_pbi": 40,
"aguardar_refresh_segundos": 30

💾 Backups automáticos
Com fazer_backup: true, uma cópia do Excel é salva em /backups/ antes de cada atualização.
Arquivos nomeados como: planilha_backup_20250311_083000.xlsx
Para desativar: mude para false no config.json.

9. Solução de problemas
!
Power BI não abre
Verifique o powerbi_exe_path no config.json. No PowerShell, rode: Get-Command PBIDesktop para achar o caminho correto.
!
Excel não detectado
Confirme que o Google Drive Desktop está rodando e sincronizou o arquivo. O caminho no config deve ser o caminho local, não o link do Drive.
!
Clique não funciona
Aumente aguardar_carregamento_pbi. Se persistir, mude auto_click.ativo para false.
!
Atualiza múltiplas vezes
Aumente cooldown_segundos (padrão: 30). O Google Drive pode salvar o arquivo em etapas.
!
Erro de permissão
Execute o iniciar.bat como Administrador (botão direito > Executar como administrador).
!
Python não encontrado
Reinstale o Python marcando 'Add Python to PATH' durante a instalação.

Automação Power BI + Excel v2.0  •  Google Drive  •  Gerado automaticamente
