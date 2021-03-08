
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DocXReader2 {

    private static final Pattern PATTERN_EXTRACT_COLON_NOT_PRECEEDED_BY = Pattern.compile("(?<!(example|Example)):\\W"); // Matches : that is not preceeded by 'example' or 'Example'. Also it requires a whitespace (\W) afterwards. This way we don't match other markups (e.g. write:ADV.PART)

    public static void main(String[] args) throws Exception{

        File file = new File("Croft-2021-Glossar_NEW_2.docx");
        FileInputStream fis = new FileInputStream(file.getAbsolutePath());

        XWPFDocument document = new XWPFDocument(fis);
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        System.out.println(paragraphs.size());

        // Structure to merge several XWPFRuns of the XWPFParagraphs into one list
        class MergedXWPFParagrahps {
            private List<XWPFRun> mergedRuns = new ArrayList<>();
            private void addXWPFRuns(List<XWPFRun> runs){
                this.mergedRuns.addAll(runs);
            }
            private List<XWPFRun> getMergedRuns(){
                return mergedRuns;
            }
        }
        List<MergedXWPFParagrahps> mergedParagraphs = new ArrayList<>();
        MergedXWPFParagrahps currentParagraph = null;



        // Scan all paragraphs
        for (XWPFParagraph para : paragraphs) {

            String paraText = para.getParagraphText();

            // Hack 0: Skip empty paragraphs
            if(paraText.trim().isEmpty()) continue;

            // HACK 1: Check if the sentence contains a ':'-character. If so, we assume that this denotes a new construction entry. (CAVE: Phrases that contain "Example:" will break this!)

            boolean newParagraph = PATTERN_EXTRACT_COLON_NOT_PRECEEDED_BY.matcher(paraText).find(); // Checks if the sentence contains a colon that is not preceeded by 'example' or 'Example'

            // HACK 2: There are some entries of the form 'aspectual structure <em>see</em> aspect'. Lets check the runs if there is an italic 'see' in there!

            if(paraText.contains("see")){
                for(XWPFRun run : para.getRuns()){
                    if(run.isItalic() && run.text().trim().equals("see")){
                        newParagraph = true;
                    }
                }
            }

            /*
            System.out.println(paraText);
            System.out.println("  getSpacingBefore: " + para.getSpacingBefore());
            System.out.println("  getIndentationFirstLine: " + para.getIndentationFirstLine());
            System.out.println("  getIndentationHanging: " + para.getIndentationHanging());
            System.out.println("  getIndentFromLeft: " + para.getIndentFromLeft());
            System.out.println("  getIndentationLeft: " + para.getIndentationLeft());
            System.out.println("  getIndentFromRight: " + para.getIndentFromRight());
            System.out.println("  getIndentationRight: " + para.getIndentationRight());
             */

            // Create a new merged paragraph object and add it to our merged paragraph list
            if(newParagraph){
                currentParagraph = new MergedXWPFParagrahps();
                mergedParagraphs.add(currentParagraph);
            }

            // If we do not have a paragraph yet -> skip until we have one!
            if(currentParagraph==null) continue;

            // Add the runs to the paragraph
            currentParagraph.addXWPFRuns(para.getRuns());
        }
        fis.close();


        XSSFWorkbook wb = new XSSFWorkbook();
        CustomSheet sem = new CustomSheet(wb.createSheet("sem"));
        CustomSheet inf = new CustomSheet(wb.createSheet("inf"));
        CustomSheet cxn = new CustomSheet(wb.createSheet("cxn"));
        CustomSheet str = new CustomSheet(wb.createSheet("str"));
        CustomSheet other = new CustomSheet(wb.createSheet("other"));

        System.out.println(mergedParagraphs.size());
        for(MergedXWPFParagrahps mergedParagraph : mergedParagraphs){
            StringBuilder sent = new StringBuilder();
            Set<String> types;
            for(XWPFRun run : mergedParagraph.getMergedRuns()){
                String text = run.getText(0);


                /*if(run.isBold()){
                    text = "<strong>" + text + "</strong>";
                }
                if(run.isItalic()){
                    text = "<em>" + text + "</em>";
                }
                if(text!=null){
                    System.out.print(text);
                }*/
                sent.append(run.toString());
            }
            types = getTypes(sent.toString());

            if(types.contains("sem")) writeConstruction(sem,  sent.toString());
            if(types.contains("inf")) writeConstruction(inf, sent.toString());
            if(types.contains("cxn")) writeConstruction(cxn, sent.toString());
            if(types.contains("str")) writeConstruction(str, sent.toString());
            if(!types.contains("sem") && !types.contains("inf") && !types.contains("cxn") && !types.contains("str")) writeConstruction(other, sent.toString());




        }
        OutputStream fileOut = new FileOutputStream("Constructions.xlsx");
        wb.write(fileOut);
        wb.close();
    }




    private static final Pattern PATTERN_EXTRACT_PARENTHESIS = Pattern.compile("\\Q(\\E([^)]*?)\\Q)\\E");
    private static Set<String> getTypes(String sent){
        Set<String> types = new TreeSet<>();
        String prefix = sent.contains(":") ? sent.substring(0, sent.indexOf(":")) : sent;
        Matcher matcher = PATTERN_EXTRACT_PARENTHESIS.matcher(prefix);
        while(matcher.find()){
            String content = matcher.group(1);
            if(!content.contains("/")){
                types.add(content);
            }
            else{
                String[] values = content.split("/");
                for(String v : values){
                    types.add(v);
                }
            }
        }
        return types;
    }

    private static  void writeConstruction(CustomSheet sheet, String sent) {
        String prefix = sent.contains(":") ? sent.substring(0, sent.indexOf(":")) : null;
        String suffix = sent.contains(":") ? sent.substring(sent.indexOf(":") + 1, sent.length()) : null;

        if(prefix == null || suffix == null) return;

        Row newRow = sheet.createRow();
        Cell firstCell = newRow.createCell(0);
        Cell secondCell = newRow.createCell(1);

        firstCell.setCellValue(prefix);
        secondCell.setCellValue(suffix);


    }

}
