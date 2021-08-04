package rmt.search;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

public class SearchingDouble {

    static String [] header = {"ФИО", "Дата рождения", "Дата начала выплаты", "Дата окончания выплаты", "Подразделение"};
    static String resultFile = "d:" + File.separator + "Результат поиска.xls";

    static SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");

    // Параметры для ФИО
    static int oneFIOCol = 0, twoFIOCol = 0;
    static int oneFIOBegin = 0, oneFIOEnd = 2,
            twoFIOBegin = 0, twoFIOEnd = 2;
    public static boolean oneUnitedFIO = true, twoUnitedFIO = false;

    // Параметры для даты рождения
    static int oneBDCol = 1, twoBDCol = 3;
    static boolean isDateCompare = true;

    static int oneRN = 2, twoRN = 0;

    public static void main(String[] args) {
        ExcelManager em = ExcelManager.getExcelManager();
        em.createNewBook();
        em.createSheet("Список");
        em.setHeader(header);

        startChecking(em);
    }

    private static void startChecking(ExcelManager em) {
        File firstFolder = new File(System.getProperty("user.dir") + "\\first\\");
        File secondFolder = new File(System.getProperty("user.dir") + "\\second\\");


        for (File firstFile : firstFolder.listFiles()) {
            Sheet firstSheet = em.getSheet(firstFile);


            for (int firstRowNumber = oneRN; firstRowNumber < firstSheet.getPhysicalNumberOfRows(); firstRowNumber++) {
                Row oneRow = firstSheet.getRow(firstRowNumber);
                System.out.println(firstRowNumber);

                String oneFIO;
                if (oneUnitedFIO) {
                    oneFIO = getFIO(oneRow.getCell(oneFIOCol));
                } else {
                    oneFIO = getFIO(oneRow, oneFIOBegin, oneFIOEnd);
                }

                String oneBD = "";
                if (isDateCompare) {
                    oneBD = getDate(oneRow, oneBDCol);
                }

                for (File secondFile : secondFolder.listFiles()) {
                    Sheet secondSheet = em.getSheet(secondFile);

                    for (int secondRowNumber = twoRN; secondRowNumber < secondSheet.getPhysicalNumberOfRows(); secondRowNumber++) {
                        Row twoRow = secondSheet.getRow(secondRowNumber);

                        String twoFIO;
                        if (twoUnitedFIO) {
                            twoFIO = getFIO(twoRow.getCell(twoFIOCol));
                        } else {
                            twoFIO = getFIO(twoRow, twoFIOBegin, twoFIOEnd);
                        }

                        String twoBD = "";
                        if (isDateCompare) {
                            twoBD = getDate(twoRow, twoBDCol);
                        }


                        if (isDateCompare) {
                            if (oneFIO.equals(twoFIO) && oneBD.equals(twoBD)) {
                                List<String> value = new ArrayList<>();
                                value.add(firstUpperCase(oneFIO));      // ФИО
                                value.add(oneBD);                       // Дата рождения
                                value.add(oneRow.getCell(2).getStringCellValue());
                                value.add(oneRow.getCell(3).getStringCellValue());
                                value.add(oneRow.getCell(4).getStringCellValue());
                                em.write(value);
                                em.save(resultFile);
                            }
                        }
                    }
                }
            }
        }

        em.closeBook();
    }

    private static String getFIO(Cell cell) {
        return cell.getStringCellValue().trim().toLowerCase();
    }

    private static String getFIO(Row row, int begin, int end) {
        StringBuilder fio = new StringBuilder();
        for (int i = begin; i <= end; i++){
            Cell cell = row.getCell(i);
            fio.append(cell.getStringCellValue().trim().toLowerCase());
            if (i != end){
                fio.append(" ");
            }
        }
        return fio.toString();
    }

    private static String getDate(Row row, int col) {
        String bd = "";
        if (row.getCell(col).getCellType() == Cell.CELL_TYPE_NUMERIC) {
            Date date = row.getCell(col).getDateCellValue();
            if (date != null) {
                bd = dateFormat.format(date);
            }
        } else {
            bd = row.getCell(col).getStringCellValue().trim();
        }

        return bd;
    }

    private static String firstUpperCase(String word) {
        if (word.contains(" ")) {
            String[] array = word.split(" ");
            StringBuilder wordBuilder = new StringBuilder();
            for (String s : array) {
                wordBuilder.append(s.substring(0, 1).toUpperCase());
                wordBuilder.append(s.substring(1));
                wordBuilder.append(" ");
            }
            word = wordBuilder.toString().trim();
        } else {
            word = word.substring(0, 1) + word.substring(1).trim();
        }

        return word;
    }
}