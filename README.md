# FriPremium — Sistema de Chamados

Abertura de chamados sem login, consulta por protocolo e área restrita (painel e gestão dos chamados).

## Arquitetura
- **Frontend:** `index.html` + `logo.png`, publicados no GitHub Pages.
- **Backend:** `apps-script/Code.gs`, vinculado à planilha "FriPremium — Chamados".
- **Fotos:** pasta privada "FriPremium Chamados - Fotos" no Drive da conta que implanta o script.

## Implantação

### 1. Planilha e Apps Script
1. Criar uma planilha Google chamada **FriPremium — Chamados**.
2. Extensões → Apps Script. Substituir o conteúdo de `Código.gs` pelo de `apps-script/Code.gs`.
3. Em Configurações do projeto, marcar "Mostrar arquivo de manifesto" e substituir `appsscript.json` pelo deste repositório.
4. Selecionar a função `setup` e clicar em Executar. Autorizar o acesso (Planilhas e Drive).
5. Implantar → Nova implantação → tipo **App da Web**:
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
6. Copiar a URL que termina em `/exec`.

### 2. Frontend
1. Em `index.html`, substituir `COLE_AQUI_A_URL_DO_APPS_SCRIPT` pela URL `/exec`.
2. Publicar `index.html` e `logo.png` no repositório `fripremium-chamados`.
3. Settings → Pages → Deploy from branch → `main` / root.

### 3. Primeiro acesso
Jeferson e Lucas entram em "Acesso restrito" com o CPF e definem a senha de 4 dígitos. Recomenda-se fazer isso no dia da implantação, porque até a senha ser definida, quem souber o CPF pode defini-la.

## Administração pela planilha
| Aba | Uso |
|---|---|
| `Chamados` | Base de chamados. Não editar manualmente as colunas de data e status. |
| `Historico` | Registro de cada mudança de status. |
| `Usuarios` | CPF, perfil (`gestor` edita, `diretor` só visualiza), área (`*` = todas). |
| `Setores` | Lista de setores e locais. Coluna `Ativo` = `Não` oculta o item. |
| `AreasDestino` | Áreas que recebem chamados e prefixo do protocolo (Manutenção = `MAN`). |
| `Categorias` | Categorias por área destino. |
| `Config` | Prazos-alvo (SLA) em horas por urgência. |

**Nova área (ex.: TI):** incluir linha em `AreasDestino` (ex.: `TI`, `TI`, `Sim`), as categorias em `Categorias` e o responsável em `Usuarios` com a área `TI`. O formulário passa a exibir o campo "Área responsável" automaticamente.

**Esqueceu a senha:** no editor do Apps Script, abrir a função `resetarSenha`, informar o CPF e executar.

**Alterações no código do Apps Script** exigem Implantar → Gerenciar implantações → editar → Nova versão, para manter a mesma URL.
