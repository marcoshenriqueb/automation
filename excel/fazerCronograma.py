import DeitaCronograma, FormataCronograma

if __name__ == '__main__':
    workbook = input('Qual o nome da planilha? ')
    margin = input('Qual a margem? ') or 0.7

    with DeitaCronograma.DeitaCronograma(workbook, margin) as trab:
        trab.transform()

    with FormataCronograma.FormataCronograma(workbook) as form:
        form.style()
