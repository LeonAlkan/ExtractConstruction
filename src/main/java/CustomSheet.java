import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class CustomSheet {
    private Sheet sheet;
    private int rowCount = 0;

    public CustomSheet (Sheet sheet) {
        this.sheet = sheet;
    }

    public Row createRow() {
        Row newRow = sheet.createRow(rowCount);
        rowCount++;

        return newRow;
    }

    public Cell createCell(Row row, int index) {
        Cell newCell = row.createCell(index);

        return newCell;
    }
}
