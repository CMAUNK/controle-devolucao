# Preparar planilhas de manifestos

Aplicativo Streamlit para ler vários manifestos Jadlog em PDF e preparar linhas separadas para as bases RECEBENDO e EXPEDINDO.

## Iniciar no Windows

1. Instale Python 3.10 ou mais recente.
2. Na primeira vez, abra o PowerShell nesta pasta e instale as dependências:

   ```powershell
   py -m pip install -r requirements.txt
   ```

3. Depois, inicie o aplicativo com duplo clique em `iniciar_app.bat`. Abra o navegador manualmente e acesse `http://localhost:8501`. Também pode iniciá-lo pelo PowerShell:

   ```powershell
   py -m streamlit run app.py --server.headless=true --server.port=8501
   ```

O aplicativo não abre o navegador sozinho. Acesse `http://localhost:8501` manualmente. Mantenha a janela preta aberta enquanto estiver usando o app; ela executa o servidor. Se o aplicativo encerrar, o iniciador mantém a janela aberta para mostrar qualquer mensagem de erro.

## Regras aplicadas

- Um manifesto cujo DESTINO contém “Rua Comandante Vergueiro da Cruz” vai para BASE-RECEBENDO. O Ponto Destino recebe “FL RIO DEVOLUÇÃO”; o Ponto Origem é procurado na tabela de endereços.
- Se o endereço aparecer em ORIGEM, o manifesto vai para BASE-EXPEDINDO. O Ponto Origem recebe “FL RIO DEVOLUÇÃO”; o Ponto Destino é procurado na tabela.
- O primeiro número de 14 dígitos de cada linha da coluna REMESSA é usado como CT-e.
- O aplicativo soma os totais de VOLUMES informados em cada página do mesmo PDF. Quando uma página não traz total impresso, usa o total indicado nas linhas dessa página.
- A quantidade de CT-e encontrados preenche a coluna de volumes não recebidos e os números são reunidos na coluna de CT-e faltantes/não recebidos.
- Status fixo RECEBENDO: “PENDENTE DE RECEBIMENTO” na coluna STATUS CTE'S FALTANTE.
- Status fixo EXPEDINDO: “NO PISO DA FL”.
- Os PDFs são processados independentemente. Não há cruzamento de informações entre arquivos.
- Endereço sem correspondência fica em branco e aparece em Avisos para revisar. Se o endereço de referência não permitir classificar o manifesto, esse PDF não gera linha.

## Tabela de endereços

O arquivo `mapeamento_enderecos.csv` contém os mapeamentos da planilha enviada, incluindo as correções combinadas:

- SÃO PAULO → TC REVERSAS-JAD-SAO
- RUA DOUTOR FREIRE CISNEIRO → FREGUESIA
- RUA COMANDANTE VERGUEIRO DA CRUZ → FL RIO DEVOLUÇÃO

No aplicativo, abra “Conferir ou ajustar a tabela de endereços” para revisar ou incluir correspondências antes de ler os PDFs. Endereços mais específicos têm prioridade sobre nomes de cidade. A tabela pode ser editada no próprio app; alterações feitas ali valem para o processamento atual.

## Resultados

- `manifestos_processados.xlsx`: duas abas, com os mesmos cabeçalhos das bases e uma linha por PDF classificado.
- `base_recebendo.tsv` e `base_expedindo.tsv`: arquivos tabulados para copiar e colar nas abas online.

As colunas definidas como ignoradas ficam vazias. Confira as linhas e os avisos antes de copiá-las para a planilha online. O aplicativo prepara arquivos para copiar; não altera diretamente a planilha online.

O leitor espera PDFs com texto selecionável, no formato dos manifestos de exemplo. PDFs digitalizados como imagem, protegidos por senha ou em layout diferente podem exigir revisão manual.

## Publicar e hospedar

Para abrir o app sem deixar um CMD rodando no notebook, o código pode ser hospedado no Streamlit Community Cloud a partir deste repositório GitHub. Depois da implantação, cada computador abre o endereço HTTPS do app no navegador.

1. Envie estes arquivos a um repositório GitHub (recomendado: privado).
2. Conecte a conta GitHub ao Streamlit Community Cloud.
3. Crie o app selecionando o repositório, a branch `main` e o arquivo `app.py`.

**Privacidade:** com o app hospedado na nuvem, os PDFs enviados pelo navegador serão processados no serviço de hospedagem. Confirme que a política da empresa permite esse uso antes de carregar manifestos reais. Se os dados não puderem sair da rede corporativa, a hospedagem deve ser em um servidor aprovado pela empresa.
