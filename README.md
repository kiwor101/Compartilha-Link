# Link Facil OneDrive

Modelo inicial de app para usuários internos enviarem arquivos para o OneDrive corporativo e receberem um link externo pronto para copiar.

Este repositório tem duas opções:

- `static/`: protótipo sem instalação, bom para testar a ideia rapidamente.
- `src/`: versão React/TypeScript para virar produto com build, testes e deploy depois.
- `.github/workflows/pages.yml`: publicação automática da pasta `static` no GitHub Pages.

## O que este modelo faz

- Login com conta Microsoft corporativa.
- Upload de um ou mais arquivos para o OneDrive do usuário logado.
- Organização automática por pasta, setor e data.
- Criação de link `Qualquer pessoa com o link pode visualizar`.
- Botão para copiar o link gerado.

## Configuração no Microsoft Entra ID

1. Acesse o Microsoft Entra admin center.
2. Vá em **App registrations** e crie um novo registro.
3. Em **Supported account types**, escolha apenas contas deste diretório organizacional.
4. Em **Redirect URI**, selecione **Single-page application (SPA)** e informe:

   ```text
   http://localhost:5173
   ```

   Para testar a versão estática deste protótipo, também cadastre:

   ```text
   http://localhost:8080
   ```

5. Copie o **Application (client) ID**.
6. Copie o **Directory (tenant) ID**.
7. Em **API permissions**, adicione permissões delegadas do Microsoft Graph:

   ```text
   User.Read
   Files.ReadWrite
   ```

8. Conceda consentimento administrativo, se o tenant exigir.

## Teste rápido sem instalação

Edite `static/config.js`:

```text
window.APP_CONFIG = {
  microsoftClientId: "seu-client-id",
  microsoftTenantId: "seu-tenant-id",
  uploadRootFolder: "Compartilhamentos Externos"
};
```

Depois rode um servidor local dentro da pasta `static`.

Se tiver Python instalado:

```text
python -m http.server 8080
```

Abra:

```text
http://localhost:8080
```

## Publicar pelo GitHub Pages

Depois de enviar os arquivos para o GitHub:

1. Abra o repositório no GitHub.
2. Vá em **Settings**.
3. Vá em **Pages**.
4. Em **Build and deployment**, selecione:

   ```text
   Source: GitHub Actions
   ```

5. Vá em **Actions** e rode o workflow **Deploy GitHub Pages**, se ele ainda não tiver rodado sozinho.

A URL esperada será parecida com:

```text
https://kiwor101.github.io/Compartilha-Link/
```

Cadastre essa URL no Microsoft Entra ID como Redirect URI de **Single-page application (SPA)**.

## Versão React

Crie um arquivo `.env` a partir de `.env.example`:

```text
VITE_MS_CLIENT_ID=seu-client-id
VITE_MS_TENANT_ID=seu-tenant-id
VITE_UPLOAD_ROOT_FOLDER=Compartilhamentos Externos
```

Depois instale e rode:

```text
npm install
npm run dev
```

Abra:

```text
http://localhost:5173
```

## Observações importantes

- Este primeiro modelo usa upload simples do Microsoft Graph, limitado a 250 MB por arquivo.
- Para arquivos maiores, a próxima evolução é usar `createUploadSession`.
- O link externo depende da política do OneDrive/SharePoint da organização. Se links anônimos estiverem bloqueados no tenant, a Microsoft recusará a criação do link.
- Para prontuários e documentos sensíveis, é recomendado adicionar expiração de link, histórico e regras por setor antes de uso em produção.
