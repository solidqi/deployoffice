# Instalação do Microsoft Office (para profissionais de TI)

Para instalar o Microsoft Office por volume, que incluem o Project e Visio, deve-se utilizar a Ferramenta de Implantação de Implantação do Office (ODT).

Os arquivos do Microsoft Office 2019 não estão disponíveis no Centro de Atendimento de Licenciamento por Volume e sim na rede de Distribuição de Conteúdo (CDN) do Office na Internet. É possível instalar o Office diretamente da CDN ou baixar os arquivos de instalação da CND para uma pasta compartilhada da sua rede. Lembre-se, em todos os métodos o ODT será utilizado.

## Ferramenta de Implantação do Office (ODT)

O primeiro passo é baixar a [Ferramenta de Implantação do Office (ODT)](https://www.microsoft.com/en-us/download/details.aspx?id=49117) e poder ser encontrada no Centro de Download da Microsoft.
Assim que baixar a Ferramenta de Implantação do Office (ODT), deve-se clicar no arquivo executável **officedeploymenttool.exe** para extraior os arquivos da ODT. Os arquivos da ODT disponibilizados pela Microsoft são: *configuration-Office365-x64.xml*, *configuration-Office365-x32.xml*, *configuration-Office2019Enterprise.xml* e *setup.exe*.
O arquivo *setup.exe* é a aplicação de linha de comando para efetuar o download e instalação do Office. Este arquivo deverá ser utilizado em conjunto com o arquivo *configuration.xml* criado e configurado por você ou utilizar os arquivos fornecidos na ferramenta ODT. É muito fácil configurar o arquivo *configuration.xml* através do Bloco de Notas ou qualquer outro editor de textos.

## Configuration.xml
Como dito anteriormente, deve-se criar o arquivo *configuration.xml*. No exemplo abaixo, foi criado o conteúdo do arquivo *configuration.xml* para instalar as aplicações Office 2019, Project 2019 e Visio 2019. Veja abaixo:

```XML
<Configuration>
  <Add SourcePath="\\<servername>\<sharedfolder>" OfficeClientEdition="64" Channel="PerpetualVL2019">
      <Product ID="ProPlus2019Volume"  PIDKEY="XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" >
         <Language ID="pt-br" />
      </Product>
      <Product ID="ProofingTools">
        <Language ID="pt-br" />
      </Product>
      <Product ID="ProjectPro2019Volume"  PIDKEY="XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" >
         <Language ID="pt-br" />
      </Product>
      <Product ID="VisioPro2019Volume"  PIDKEY="XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" >
         <Language ID="pt-br" />
      </Product>
  </Add>
<RemoveMSI />
<Display Level="None" AcceptEULA="TRUE" />  
</Configuration>
```

Muito bem, como pode ser observados temos uma série de parâmentros que compõem o conteúdo do arquivo *configuration.xml*. Os parâmetros mais importantes são:

- **SourcePath** - É pasta compartilhada onde serão salvos os arquivos de instalação do Office;
- **OfficeClientEdition** - Arquietetura do Microsoft Office, pode-se x64 ou x32;
- **ProductID** - Tipo do produto;
- **PIDKEY** - Chave de ativação do produto;
- **Language ID** - Linguagem de instalação;
- **RemoveMSI** - Remove ou não o produto caso esteja previamente instalado
