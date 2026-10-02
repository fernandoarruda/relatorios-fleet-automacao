# Relatórios Fleet para caminhonetes

Este pacote gera DOCX e PDF a partir do formato do relatório **Caminhonete CH14** quando houver Microsoft Word no Windows ou LibreOffice disponível. O modelo preserva as figuras técnicas e a paginação do relatório original. Cada TAG recebe seus próprios dados, as 24 verificações, nove evidências obrigatórias e até oito imagens adicionais na Figura 7.2.

## Preparar no Windows

1. Extraia o ZIP mantendo todas as pastas.
2. Instale Python 3, caso ainda não esteja disponível, e no terminal execute, dentro da pasta extraída:

   ```powershell
   py -3 -m pip install -r requirements.txt
   ```

3. Para exportar PDF automaticamente, tenha o Microsoft Word instalado no Windows ou o LibreOffice. O gerador tenta primeiro o Word no Windows e usa o LibreOffice como alternativa. Se a conversão falhar, informa o motivo e mantém o DOCX.

## Preencher cada ativo

1. Na aba **Ativos** de `Controle_Ativos.xlsx`, revise a linha da TAG. Campos obrigatórios: `TAG`, `MODELO`, `VERSAO_MODELO`, `DATA_INICIO`, `DATA_CONCLUSAO`, `DATA_COMISSIONAMENTO`, `DATA_EMISSAO` e `AVALIACAO_INICIAL=APROVADO`. O campo `MODELO` aparece na tabela 5.1; `VERSAO_MODELO` substitui o modelo no objetivo. Mantenha `GERAR=NAO` durante a revisão. CB26 e CB08 não têm data de comissionamento no controle recebido.
2. Na aba **Verificacoes**, os 24 itens com `TAG=PADRAO` são usados em todos os relatórios, inclusive condições e observações. Confirme se esses resultados se aplicam de fato ao ativo antes de mudar `GERAR` para `SIM`. Caso uma TAG precise de resultado ou observação próprios, acrescente uma linha com a TAG real e o mesmo `GRUPO` e `NUMERO`; ela substitui apenas o item correspondente do padrão. O gerador só emite o documento conclusivo se as 24 condições resultantes forem `OK`.
3. Crie uma pasta `fotos/NOVA_TAG/` e coloque as nove evidências obrigatórias abaixo. A extensão pode ser `.jpg`, `.jpeg`, `.png`, `.webp`, `.tif` ou `.tiff`. Para as posições individuais, coloque exatamente um arquivo com o nome indicado. O gerador ajusta as imagens às áreas do modelo sem deformá-las.

| Arquivo sem extensão | Evidência no relatório |
|---|---|
| `foto_ativo` | Figura 2.1, caminhonete |
| `foto_antes` | Figura 7.1, ativo como recebido, antes da intervenção |
| `foto_antes_adicional1` a `foto_antes_adicional8` | Figura 7.2, imagens adicionais do ativo como recebido; opcionais, até oito |
| `foto_painel` | Figura 9.2, painel instalado |
| `foto_gnss` | Figura 9.3, antena GNSS |
| `foto_chicotes1` a `foto_chicotes4` | Figura 9.4, chicotes instalados; use de uma a quatro fotos |
| `print_classe` | Figura 10.1, classe no MineStar |
| `print_configuracao` | Figura 10.2, configuração embarcada |
| `foto_comissionado` | Figura 10.3, equipamento comissionado |
| `checklist_assinado` | Figura 11.1, checklist assinado em imagem; alternativa: PDF por TAG |

**Mosaico da Figura 9.4:** nomeie os arquivos `foto_chicotes1.jpg`, `foto_chicotes2.png` etc., com os números **1 a 4 imediatamente antes da extensão**. Pode fornecer qualquer quantidade de uma a quatro fotos. O gerador as ordena pelo número e monta o mosaico no espaço existente, preservando as proporções e exibindo a imagem inteira, sem cortes. Não use dois arquivos para o mesmo número. Para manter a compatibilidade com pastas antigas, `foto_chicotes.png` continua aceito quando não houver nenhuma foto numerada; não misture os dois padrões.

**Mosaico da Figura 7.2:** use `foto_antes_adicional1.jpg`, `foto_antes_adicional2.png` etc. na pasta `fotos/TAG/`. Os números aceitos são **1 a 8**, imediatamente antes da extensão, e podem ter lacunas (por exemplo, 1, 3 e 5). O gerador ordena as fotos, adapta a altura do mosaico à quantidade e mostra cada imagem inteira, sem cortes. Não repita um número em extensões diferentes. Se nenhuma foto adicional existir, a figura e a legenda 7.2 serão omitidas no relatório gerado; `foto_antes` da Figura 7.1 continua obrigatória.

**Checklist em PDF:** crie a pasta `checklists/` ao lado de `gerar_relatorios.py` e coloque, por exemplo, `CD128 - Checklist comissionamento.pdf`. O nome deve começar pela TAG seguida de separador e da palavra `Checklist`. Também é aceito na pasta `fotos/CD128/`. O PDF tem prioridade sobre `checklist_assinado.png` quando ambos existirem. Todas as páginas do PDF entram no item 11.1, cada uma em uma página do relatório, e a legenda fica após a última. O gerador bloqueia a TAG se encontrar mais de um PDF correspondente. Para usar outra pasta, passe `--checklists "C:\caminho\da\pasta"`. Instale a dependência nova com `py -3 -m pip install -r requirements.txt`.

4. Confira as datas e as fotos, depois altere `GERAR` para `SIM`. **Não reutilize a pasta de fotos do CH14** para outro ativo.

Na aba **Ativos**, as colunas L a AE acompanham as evidências: ✓ verde quando o arquivo é encontrado, ✗ vermelho quando falta uma evidência obrigatória e ! vermelho para arquivos conflitantes. As fotos de chicotes 2 a 4 e as fotos adicionais 7.2 são opcionais: suas células vazias ficam sem cor. Feche a planilha no Excel e execute `ATUALIZAR_EVIDENCIAS.bat` para atualizar os checks sem gerar relatórios. `GERAR_RELATORIOS.bat` executa a atualização antes da geração. O painel confere a existência dos arquivos, não o conteúdo ou a qualidade das fotos.

## Gerar os arquivos

Com o terminal aberto na pasta extraída:

```powershell
py -3 gerar_relatorios.py
```

Ou dê dois cliques em `GERAR_RELATORIOS.bat`. Os arquivos aparecem em `saida/`, com os nomes `Checklist_Entrega_TAG.docx` e `Checklist_Entrega_TAG.pdf` quando a exportação por Word ou LibreOffice funcionar. Para apenas uma TAG, use `py -3 gerar_relatorios.py --tag CH15`. O gerador relata quais ativos foram bloqueados e por quê; um erro em uma TAG não interrompe as demais.

Para reproduzir o exemplo CH14 em **rascunho**, execute:

```powershell
py -3 gerar_relatorios.py --teste --tag CH14
```

**Atenção ao CH14:** `fotos/CH14/foto_antes.png` é apenas uma imagem de teste com aviso visível; substitua pela foto real do ativo como recebido antes de emitir como `SIM`. O controle de tripulados registra Ford Ranger, instalação de 06 a 07/08/2026 e comissionamento em 07/08/2026, diferindo do relatório anterior. Confirme os dados reais antes de emitir. A data de emissão de CH14 (13/08/2026) veio do quadro de revisões.

Antes de enviar um relatório, confira no Word/PDF: TAG do cabeçalho, datas, revisão, imagens, resultados, legendas e paginação. A primeira versão foi preparada para **caminhonetes com o mesmo escopo e o mesmo conjunto de itens** do CH14. Casos com checklist diferente ou pendências exigem adaptar o modelo e as conclusões.
