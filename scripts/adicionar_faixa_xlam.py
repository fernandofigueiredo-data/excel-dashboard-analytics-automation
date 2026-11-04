#!/usr/bin/env python3
"""Adiciona a guia "Dashboard" na Faixa de Opcoes de um suplemento .xlam.

Uso:  python3 adicionar_faixa_xlam.py DashboardGenerator.xlam
Gera: DashboardGenerator_faixa.xlam (o original nao e alterado).

O Excel nao cria a Faixa de Opcoes pelo VBA; ela vem de um arquivo XML
(customUI) guardado dentro do .xlam, que e um ZIP. Este script so acrescenta
esse XML e a relacao que o aponta; nao mexe no projeto VBA.
Os botoes chamam modDashboard.Faixa_AtualizarTudo, Faixa_AplicarTema e
modFixar.Faixa_FixarGrafico.
Feche o Excel antes de rodar.
"""
import sys
import zipfile
from pathlib import Path

CUSTOM_UI = """<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="dsTab" label="Dashboard">
        <group id="dsGrupo" label="Dashboard">
          <button id="dsAtualizar" label="Atualizar dashboard" size="large"
                  imageMso="RefreshAll" onAction="Faixa_AtualizarTudo"/>
          <button id="dsTema" label="Aplicar tema" size="large"
                  imageMso="ThemesGallery" onAction="Faixa_AplicarTema"/>
          <button id="dsFixar" label="Fixar gr&#225;fico" size="large"
                  imageMso="Pushpin" onAction="Faixa_FixarGrafico"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""
REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
REL = f'<Relationship Id="rIdDashboardUI" Type="{REL_TYPE}" Target="customUI/customUI.xml"/>'


def main() -> None:
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    origem = Path(sys.argv[1])
    if origem.suffix.lower() != ".xlam" or not origem.is_file():
        sys.exit(f"Arquivo .xlam nao encontrado: {origem}")
    destino = origem.with_name(origem.stem + "_faixa.xlam")

    with zipfile.ZipFile(origem) as zin, zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
        nomes = zin.namelist()
        if "customUI/customUI.xml" in nomes or "customUI/customUI14.xml" in nomes:
            sys.exit("Este .xlam ja tem Faixa de Opcoes personalizada; nada foi alterado.")
        for item in zin.infolist():
            dados = zin.read(item.filename)
            if item.filename == "_rels/.rels":
                texto = dados.decode("utf-8")
                texto = texto.replace("</Relationships>", REL + "</Relationships>")
                dados = texto.encode("utf-8")
            elif item.filename == "[Content_Types].xml":
                texto = dados.decode("utf-8")
                if 'Extension="xml"' not in texto:
                    fim_tag = texto.index(">", texto.index("<Types")) + 1
                    texto = (texto[:fim_tag]
                             + '<Default Extension="xml" ContentType="application/xml"/>'
                             + texto[fim_tag:])
                dados = texto.encode("utf-8")
            zout.writestr(item, dados)
        zout.writestr("customUI/customUI.xml", CUSTOM_UI)
    print(f"Criado: {destino}")


if __name__ == "__main__":
    main()
