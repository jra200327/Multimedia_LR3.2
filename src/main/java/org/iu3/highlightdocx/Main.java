package org.iu3.highlight;

/**
 * Консольное приложение.
 * Аргументы командной строки:
 *   args[0] — путь к исходному DOCX
 *   args[1] — искомое слово
 *   args[2] — путь к выходному DOCX
 *
 * Пример:
 *   java -jar HighlightDocx.jar input.docx "ве" output.docx
 */
public class Main {

    public static void main(String[] args) {
        if (args.length < 3) {
            System.out.println("Использование: java -jar HighlightDocx.jar " +
                    "<входной.docx> <слово> <выходной.docx>");
            return;
        }

        String inputPath  = args[0];
        String searchWord = args[1];
        String outputPath = args[2];

        try {
            DocxHighlighter.highlight(inputPath, searchWord, outputPath);
            System.out.println("Готово: " + outputPath);
        } catch (Exception e) {
            System.err.println("Ошибка: " + e.getMessage());
            e.printStackTrace();
        }
    }
}