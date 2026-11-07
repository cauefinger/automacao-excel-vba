# Controle de prazos — macro no LibreOffice Calc

Aprenda VBA/Basic automatizando um controle de prazos: você seleciona uma pessoa na aba **Pessoas**, roda a macro, e ela vai para a aba **Prazos** com o prazo limite calculado e colorido.

## Como usar
1. Abra `controle_prazos.ods` (a macro já está dentro). Se aparecer aviso de macros: *Ferramentas → Opções → LibreOffice → Segurança → Segurança de macros → Média*, e reabra o arquivo clicando em **Ativar macros**.
2. Na aba **Pessoas**, clique em qualquer célula de uma linha com dados.
3. *Ferramentas → Macros → Executar Macro → controle_prazos.ods → Standard → Module1 → `EnviarParaPrazo`*.
4. Rode de novo na mesma linha: a macro avisa que já foi enviada (anti-duplicata).

### Botão (opcional)
*Exibir → Barras de ferramentas → Controles de formulário*, ative o **Modo de design**, insira um **Botão**, clique com o direito → *Controle → Eventos → Ao clicar → Macro...* e escolha `EnviarParaPrazo`. Desative o modo de design para usar.

## Código: `src/PrazosModulo.bas`
| Etapa | Macro | O que você aprende |
|---|---|---|
| 1 | `OlaMacro` | variáveis, `MsgBox`, linha selecionada |
| 2 | `EnviarParaPrazo` | `If`, validação, ler/gravar células, formatar data, funções |
| 3 | `AtualizarSituacao` | laço `For`, comparar datas, cores |

## VBA do Excel × Basic do LibreOffice
| Excel/VBA | LibreOffice Basic (este projeto) |
|---|---|
| `ws.Cells(2, 1)` (base 1) | `ws.getCellByPosition(0, 1)` (base 0: coluna, linha) |
| `.Value` de texto | `.String`; número/data: `.Value` |
| `ActiveCell.Row` | `ThisComponent.CurrentSelection.getRangeAddress().StartRow + 1` |
| `IsDate(x)` | `celula.Type = com.sun.star.table.CellContentType.VALUE` |
| `.Interior.Color` | `.CellBackColor` |

Só `Dim`, `If`, `For`, `Do While`, `Sub`, `Function`, `Const`, `MsgBox` são praticamente idênticos.

## Exercícios
1. Mude `PRAZO_DIAS` para 10 e envie uma nova pessoa.
2. Mude `DIAS_ALERTA` para 4 e veja as cores mudarem (`AtualizarSituacao`).
3. Adicione uma coluna "Setor" em Pessoas e copie-a para Prazos.
4. Crie uma macro `LimparPrazos` que apague as linhas de Prazos (dica: laço `For` + `.String = ""`).
5. Faça `EnviarParaPrazo` aceitar várias linhas selecionadas.

## Arquivos
- `controle_prazos.ods` — planilha com dados fictícios e macro embutida
- `src/PrazosModulo.bas` — código-fonte da macro (texto para ler/copiar)
- `src/gerar_planilha.py`, `src/embutir_e_testar.py` — scripts que geraram/testaram a planilha (opcionais)
