# Checklist de Inspeção – Empilhadeira / Reach Stacker (Rev. 07)

Site estático (GitHub Pages) + servidor em Google Apps Script.

- `index.html` – o checklist (mesmas perguntas do formulário de papel)
- `config.js` – onde se cola a URL do servidor
- `servidor/Code.gs` – código do servidor (guarda os checklists e envia os e-mails)

## Publicar
1. **Servidor:** em script.google.com, novo projeto, cole `servidor/Code.gs`, ajuste `DESTINATARIOS`, execute `autorizar`, depois *Implantar > Nova implantação > Aplicativo da Web* (Executar como: Eu; Acesso: Qualquer pessoa). Copie a URL.
2. **config.js:** cole a URL no lugar de `COLE_AQUI_A_URL_DO_APPS_SCRIPT`.
3. **Site:** Settings > Pages > Deploy from a branch > `main` / `(root)`.
4. **QR Code de cada máquina:** `https://SEU-USUARIO.github.io/NOME-DO-REPOSITORIO/#RS6`
