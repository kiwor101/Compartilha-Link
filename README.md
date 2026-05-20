# Compartilha Link

Aplicativo interno para enviar documentos ao OneDrive corporativo e gerar links externos de forma simples.

O objetivo e facilitar o dia a dia de setores que precisam compartilhar prontuarios, documentos de prestacao de contas, imagens e arquivos grandes sem depender da interface completa do OneDrive.

## O que o app faz

- Login com conta Microsoft corporativa do dominio `@santacasaandradina.org`.
- Envio de varios arquivos de uma unica vez.
- Suporte a arrastar arquivos ou pastas inteiras para a tela.
- Upload de arquivos pequenos e grandes usando Microsoft Graph.
- Criacao de pasta dentro de `Compartilhamentos Externos`.
- Geracao de link externo com validade de 7, 30, 60 ou 90 dias.
- Historico por usuario salvo no proprio OneDrive.
- Busca no historico por nome da pasta ou data.
- Renovacao de links vencidos.
- Pagina administrativa para listar e apagar pastas antigas com base no historico.

## Enderecos

Aplicativo principal:

```text
https://kiwor101.github.io/Compartilha-Link/
```

Pagina administrativa:

```text
https://kiwor101.github.io/Compartilha-Link/admin.html
```

## Estrutura importante

```text
index.html          Tela principal do app
app.js              Login, upload, links e historico
admin.html          Tela administrativa
admin.js            Busca e limpeza de pastas antigas
styles.css          Interface visual
config.js           IDs do aplicativo Microsoft
static/             Pasta publicada pelo GitHub Pages
```

Importante: o GitHub Pages deste projeto publica a pasta `static/`. Sempre que alterar arquivos da raiz usados no site, mantenha a copia correspondente dentro de `static/`.

## Configuracao Microsoft Entra

O app usa Microsoft Graph com login delegado. A aplicacao registrada no Microsoft Entra precisa ter:

- Tipo: Single-page application (SPA)
- Redirect URI principal:

```text
https://kiwor101.github.io/Compartilha-Link/
```

- Redirect URI administrativa:

```text
https://kiwor101.github.io/Compartilha-Link/admin.html
```

Permissoes delegadas do Microsoft Graph:

```text
User.Read
Files.ReadWrite
```

Essas permissoes precisam estar consentidas pelo administrador.

## Configuracao do app

O arquivo `config.js` define os dados do Microsoft Entra:

```js
window.APP_CONFIG = {
  microsoftClientId: "CLIENT_ID",
  microsoftTenantId: "TENANT_ID",
  uploadRootFolder: "Compartilhamentos Externos"
};
```

## Publicacao

O deploy e feito pelo GitHub Actions em:

```text
.github/workflows/pages.yml
```

Ele publica a pasta:

```text
static
```

Depois de um push na branch `main`, o GitHub Pages pode levar alguns minutos para atualizar. Se parecer que nada mudou, abra com um parametro de cache:

```text
https://kiwor101.github.io/Compartilha-Link/?nocache=1
```

## Sugestao de endereco melhor

O endereco `kiwor101.github.io` funciona, mas nao e ideal para uso interno.

A melhor opcao sem custo adicional e criar um subdominio no dominio que a instituicao ja possui, por exemplo:

```text
compartilha.santacasaandradina.org
```

ou:

```text
links.santacasaandradina.org
```

Esse subdominio pode apontar para o GitHub Pages usando DNS. O GitHub Pages permite dominio personalizado sem cobrar por isso. O unico custo seria ter o dominio, mas nesse caso a instituicao ja usa `santacasaandradina.org`.

Depois de configurar o dominio, tambem sera necessario cadastrar a nova URL como Redirect URI no Microsoft Entra.

## Observacoes de seguranca

- Os arquivos ficam no OneDrive da conta que fez login.
- O historico tambem fica no OneDrive da conta logada.
- Links externos devem ser usados com validade curta sempre que possivel.
- Documentos sensiveis devem seguir as regras internas da instituicao.
