# Compartilha Link

Aplicativo interno da Santa Casa de Andradina para enviar arquivos ao OneDrive corporativo e gerar links externos com validade definida.

O foco do projeto e simplificar uma tarefa comum: o usuario entra com a conta Microsoft corporativa, escolhe os arquivos ou uma pasta inteira, define o nome da pasta e recebe um link pronto para copiar.

## Enderecos

Aplicativo principal:

```text
https://compartilhalink.santacasaandradina.org/
```

Endereco original do GitHub Pages:

```text
https://kiwor101.github.io/Compartilha-Link/
```

Pagina administrativa, mantida apenas para uso interno da TI:

```text
https://kiwor101.github.io/Compartilha-Link/admin.html
```

## Funcionalidades

- Login com conta Microsoft do dominio `@santacasaandradina.org`.
- Upload de varios arquivos de uma vez.
- Suporte a arquivos grandes usando sessao de upload do Microsoft Graph.
- Suporte a arrastar arquivos ou pastas inteiras para a tela.
- Criacao automatica da pasta base `Compartilhamentos Externos`.
- Criacao de uma pasta por envio, usando o nome informado pelo usuario.
- Bloqueio para evitar criar duas pastas com o mesmo nome no historico.
- Geracao de link externo com validade de 7, 30, 60 ou 90 dias.
- Validade padrao de 90 dias.
- Historico salvo no OneDrive da conta logada.
- Busca no historico por nome da pasta ou data.
- Botao de renovacao exibido apenas quando o link esta vencido.
- Limpeza da lista de arquivos apos envio concluido.
- Pagina administrativa para localizar e apagar pastas antigas.

## Como o usuario usa

1. Acessa o aplicativo principal.
2. Clica em `Login`.
3. Entra com a conta corporativa.
4. Informa o `Nome da pasta`.
5. Escolhe a validade do link.
6. Seleciona ou arrasta os arquivos.
7. Clica em `Enviar e gerar links`.
8. Copia o link gerado no historico.

O nome da pasta deve ser algo facil de localizar depois, como paciente, prontuario, processo ou prestacao de contas.

## Onde os arquivos ficam

Os arquivos ficam no OneDrive da propria conta que fez login.

Estrutura padrao:

```text
OneDrive
└── Compartilhamentos Externos
    ├── Nome da pasta criada pelo usuario
    └── _Compartilha Link Sistema
        └── historico.json
```

O arquivo `historico.json` guarda os links exibidos no app. Ele tambem fica no OneDrive do usuario logado.

## Estrutura do projeto

```text
index.html       Tela principal
app.js           Login, upload, links, validade e historico
admin.html       Tela administrativa da TI
admin.js         Busca e limpeza de pastas antigas
styles.css       Interface visual
config.js        Dados do aplicativo Microsoft
favicon.svg      Icone do site/favoritos
static/          Arquivos publicados pelo GitHub Pages
```

Importante: o GitHub Pages publica a pasta `static/`. Quando alterar `index.html`, `app.js`, `admin.html`, `admin.js`, `styles.css`, `config.js` ou `favicon.svg` na raiz, mantenha a copia correspondente em `static/`.

## Configuracao Microsoft Entra

Aplicativo registrado no Microsoft Entra:

```text
Compartilha Link
```

Tipo da plataforma:

```text
Single-page application (SPA)
```

Redirect URIs cadastradas:

```text
https://compartilhalink.santacasaandradina.org/
https://kiwor101.github.io/Compartilha-Link/
https://kiwor101.github.io/Compartilha-Link/admin.html
```

Permissoes delegadas do Microsoft Graph:

```text
User.Read
Files.ReadWrite
```

Essas permissoes precisam estar concedidas pelo administrador do locatario.

## Configuracao local

O arquivo `config.js` contem os dados usados pelo app:

```js
window.APP_CONFIG = {
  microsoftClientId: "CLIENT_ID",
  microsoftTenantId: "TENANT_ID",
  uploadRootFolder: "Compartilhamentos Externos"
};
```

No projeto atual:

```text
Client ID: 03b8c0d8-3d68-44c2-a2f6-18c09d7c3bca
Tenant ID: 251017e3-b3a8-40f0-b6fe-57b159bcc5d8
```

## Publicacao

O deploy e feito pelo GitHub Pages usando GitHub Actions.

Workflow:

```text
.github/workflows/pages.yml
```

Branch:

```text
main
```

Pasta publicada:

```text
static
```

Depois de enviar alteracoes para a branch `main`, o GitHub Pages pode levar alguns minutos para atualizar.

Para forcar teste sem cache:

```text
https://compartilhalink.santacasaandradina.org/?nocache=1
```

## Dominio

O dominio principal esta configurado como subdominio:

```text
compartilhalink.santacasaandradina.org
```

No DNS da Hostinger, o subdominio aponta para:

```text
kiwor101.github.io
```

Tipo de registro:

```text
CNAME
```

O dominio de e-mail Microsoft 365 pode continuar funcionando normalmente. O subdominio usado pelo app e separado dos registros MX, SPF, DKIM e autodiscover.

## Pagina administrativa

A pagina `admin.html` foi criada para a TI localizar pastas antigas dentro de `Compartilhamentos Externos`.

Ela permite:

- Escolher um periodo de corte de 1 a 10 anos.
- Buscar pastas antigas.
- Apagar as pastas encontradas.

Como ela age no OneDrive da conta logada, a TI deve entrar com a conta que possui os arquivos que serao analisados.

## Seguranca e cuidados

- O app nao armazena arquivos fora do Microsoft 365.
- O link externo deve ser usado somente quando necessario.
- A validade do link reduz o risco de acesso antigo continuar aberto.
- Documentos sensiveis devem seguir as regras internas da instituicao.
- Quem recebe o link pode acessar o conteudo enquanto o link estiver valido.
- O app usa permissao delegada: ele age como o usuario logado, nao como administrador global.

## Testes ja realizados

- Login no dominio principal.
- Upload real manual.
- Geracao de link externo.
- Historico por usuario.
- Busca por nome da pasta.
- Busca por data.
- Validade exibida no historico.
- Lista de arquivos limpa apos envio.
- Validacao para nao enviar sem nome de pasta.
- Checagem local de JavaScript.
- Checagem de consistencia entre raiz e `static/`.

## Observacao tecnica

O app e estatico e roda direto no navegador. A comunicacao com o Microsoft 365 e feita pelo Microsoft Graph usando OAuth/PKCE.

Nao ha backend proprio, banco de dados externo ou servidor pago.
