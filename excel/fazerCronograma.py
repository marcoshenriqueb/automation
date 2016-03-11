import DeitaCronograma, FormataCronograma

if __name__ == '__main__':
    workbook = input('Qual o nome da planilha? ')

    with DeitaCronograma.DeitaCronograma(workbook) as trab:
        trab.transform()

    with FormataCronograma.FormataCronograma(workbook) as form:
        form.style()
