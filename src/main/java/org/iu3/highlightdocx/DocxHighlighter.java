package org.iu3.highlight;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHighlight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class DocxHighlighter {

    /**
     * Читает inputPath, находит все вхождения searchWord,
     * выделяет их жёлтым цветом и сохраняет в outputPath.
     */
    public static void highlight(String inputPath,
                                 String searchWord,
                                 String outputPath) throws Exception {

        try (FileInputStream in  = new FileInputStream(inputPath);
             XWPFDocument    doc = new XWPFDocument(in)) {

            // Перебираем все параграфы документа
            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                processParagraph(paragraph, searchWord);
            }

            // Сохраняем результат
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                doc.write(out);
            }
        }
    }

    /**
     * Обрабатывает один параграф: находит раны с искомым словом
     * и заменяет их на три рана (до / выделенное слово / после).
     */
    private static void processParagraph(XWPFParagraph paragraph,
                                         String searchWord) {

        // Список ранов, которые нужно заменить
        List<XWPFRun> runsToReplace = new ArrayList<>();

        // 1. Собираем раны, содержащие искомое слово
        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.text();
            if (text != null && text.contains(searchWord)) {
                runsToReplace.add(run);
            }
        }

        // 2. Для каждого найденного рана создаём три новых
        for (XWPFRun oldRun : runsToReplace) {
            String text = oldRun.text();
            
            // Ищем позицию ДО удаления, чтобы не потерять данные
            int pos = text.indexOf(searchWord);
            if (pos == -1) continue; // на всякий случай

            String before = text.substring(0, pos);
            String match  = searchWord;
            String after  = text.substring(pos + searchWord.length());

            // Индекс старого рана в параграфе
            int index = paragraph.getRuns().indexOf(oldRun);
            
            // Удаляем старый ран
            paragraph.removeRun(index);

            // Вставляем новые раны
            if (!before.isEmpty()) {
                XWPFRun r1 = paragraph.insertNewRun(index);
                r1.setText(before);
                index++;
            }

            // Ран с выделенным словом
            XWPFRun r2 = paragraph.insertNewRun(index);
            r2.setText(match);
            setHighlight(r2);
            index++;

            if (!after.isEmpty()) {
                XWPFRun r3 = paragraph.insertNewRun(index);
                r3.setText(after);
            }
        }
    }
    
    /**
     * Устанавливает жёлтое выделение для рана.
     * В XWPFRun нет готового метода, поэтому работаем через XML (getCTR()).
     * Тег <w:highlight w:val="yellow"/> — стандартное выделение Word.
     */
    private static void setHighlight(XWPFRun run) {
        CTR ctr = run.getCTR();

        if (!ctr.isSetRPr()) {
            ctr.addNewRPr();
        }

        CTHighlight highlight = ctr.getRPr().addNewHighlight();
        highlight.setVal(STHighlightColor.YELLOW);
    }
}