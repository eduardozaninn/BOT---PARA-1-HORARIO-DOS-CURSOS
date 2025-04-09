# DisparaMensagemButton 2.0

Este é um aplicativo desenvolvido em Python, utilizando a biblioteca `Tkinter` para a interface gráfica e a biblioteca `ttkbootstrap` para temas personalizados. O objetivo deste programa é permitir o envio de mensagens automatizadas via WhatsApp para um conjunto de números de telefone extraídos de uma planilha Excel. Ele foi projetado para ser usado em campanhas de marketing e comunicação, permitindo a personalização de mensagens para diferentes cursos oferecidos pela empresa.

## Funcionalidades

- **Envio de Mensagens Automatizado**: Envia mensagens personalizadas via WhatsApp, incluindo informações sobre cursos oferecidos pela instituição parceira.
- **Histórico de Mensagens**: Registra as mensagens enviadas em um arquivo de log para acompanhamento.
- **Seleção de Curso e Parceiro**: Permite ao usuário selecionar o curso e a instituição parceira para enviar as mensagens.
- **Interface Gráfica**: Criada com Tkinter e ttkbootstrap para oferecer uma experiência interativa e moderna.
- **Suporte a Planilhas Excel**: O aplicativo lê os dados dos alunos a partir de planilhas Excel (.xlsx), extraindo informações como nome completo e número de telefone.
- **Temas Personalizados**: Permite a mudança de temas da interface para personalizar a aparência do aplicativo.

## Dependências

O projeto requer as seguintes bibliotecas Python:

- `tkinter`
- `ttkbootstrap`
- `openpyxl`
- `pandas`
- `pyautogui`
- `webbrowser`
- `logging`
- `json`
- `os`
- `time`
- `datetime`

Você pode instalar as dependências necessárias usando o seguinte comando:

```bash
pip install ttkbootstrap openpyxl pandas pyautogui


Como Usar
Clone este repositório para o seu computador.

Certifique-se de que o Python está instalado e as dependências estão satisfeitas.

Coloque sua planilha Excel (no formato .xlsx) com os dados dos alunos na mesma pasta do script.

Execute o script DisparaMensagemButton 2.0.py.

A interface gráfica será aberta. Selecione o tema, o curso e a instituição parceira.

Insira as informações adicionais, como horário, idade mínima, e a linha da planilha a partir da qual começar a enviar as mensagens.

Clique em "Enviar Mensagens" para iniciar o envio automático.

Personalização
Você pode personalizar a mensagem enviada e os cursos oferecidos alterando as seções relevantes no código, como o conteúdo da mensagem dentro da função send_messages().

Créditos
Desenvolvido por Eduardo Zanin e
Lucas Ferrari.
