import docx
from citation_parser import CitationParser
from csl_py import CSL


def convert_bibliography(docx_file, target_style):
    # Leer el archivo de Word
    document = docx.Document(docx_file)

    # Inicializar el parser de citas y el convertidor de estilos
    parser = CitationParser()
    converter = CSL(target_style)

    # Iterar a través de cada párrafo del documento
    for para in document.paragraphs:
        # Extraer información de las citas bibliográficas del párrafo
        citations = parser.extract_citations(para.text)

        # Si se encontraron citas, convertir cada una al estilo deseado
        if citations:
            for citation in citations:
                ref = converter.render(citation, 'bibliography')
                para.text = para.text.replace(citation.text, ref)

    # Guardar el documento con las citas convertidas
    document.save(docx_file)


if __name__ == '__main__':
    docx_file = 'documento.docx'
    target_style = 'apa'  # Estilo APA
    convert_bibliography(docx_file, target_style)
