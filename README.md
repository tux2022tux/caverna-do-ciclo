# A Caverna do Ciclo

Jogo de navegador para praticar o **ciclo trigonométrico**, estrelado pelo **Mather**. O aluno responde `cos θ` ou `sen θ` e o Mather anda esse valor pela caverna. Cos anda na horizontal, sen na vertical e 0 deixa o Mather parado.

**Jogar:** https://divonei.com/caverna-do-ciclo/

## Conteúdo

| Caverna | Ângulos |
|---|---|
| 1 · Entrada | 0, π/2, π, 3π/2 |
| 2 · Galeria dos Graus | 0°, 90°, 180°, 270°, 360° |
| 3 · Túnel Misto | radianos e graus juntos |
| 4 · Poço das Voltas | 2π, 5π/2, 3π, 540°, 720° … |
| 5 · Câmara Invertida | −π/2, −π, −270° … |

Modos **Calmo** (sem tempo) e **Desafio** (com tempo). Funciona no celular, com toque, mouse ou teclado (1, 2, 3, Enter, Esc).

## Resultados dos alunos

Ao final de cada partida, o jogo envia nome, turma, pontos, acertos e erros para uma planilha Google, por meio do script em [`apps-script/Code.gs`](apps-script/Code.gs). Para usar outra planilha, implante o script como *App da Web* e troque a constante `ENDPOINT` no `index.html`.

O jogo coleta só o primeiro nome e a turma.

## Licença

[CC BY-NC 4.0](LICENSE): livre para usar e adaptar em sala de aula, com crédito a Divonei · divonei.com e sem fins comerciais.
