/**Работа с Microsoft Excel на Java | Baeldung*/

package com.example.readexcelwithpoi;

import java.io.File;
import java.io.FileInputStream;

public class ReadXLWithPOI {
    /**Чтение из Excel*/
    /**
     * Сначала давайте откроем файл из заданного местоположения:
     */

    FileInputStream file = new FileInputStream(new File(/*fileLocation*/));
    Workbook workbook = new XSSFWorkbook(file);

    /**
     * Далее давайте извлекем первый лист файла и пройдемся по каждой строке:
     */
    Sheet sheet = workbook.getSheetAt(0);

    Map<Integer, List<String>> data = new HashMap<>();
    int i = 0;
for(
    Row row :sheet)

    {
        data.put(i, new ArrayList<String>());
        for (Cell cell : row) {
            switch (cell.getCellType()) {
                case STRING: ...break;
                case NUMERIC: ...break;
                case BOOLEAN: ...break;
                case FORMULA: ...break;
                default:
                    data.get(new Integer(i)).add(" ");
            }
        }
        i++;
    }
/**Apache POI имеет разные методы для чтения каждого типа данных.Давайте подробнее рассмотрим содержание каждого описанного выше случая переключения.

 Если значение перечисления типа ячейки равно STRING, содержимое будет считано с помощью метода getRichStringCellValue() интерфейса Cell:*/
data.get(new Integer(i)).add(cell.getRichStringCellValue().getString());

/**Ячейки, имеющие ЧИСЛОВОЙ тип содержимого, могут содержать либо дату, либо число и считываются следующим образом:*/
if (DateUtil.isCellDateFormatted(cell)) {
        data.get(i).add(cell.getDateCellValue() + "");
    } else {
        data.get(i).add(cell.getNumericCellValue() + "");
    }

/**Для логических значений у нас есть метод getBooleanCellValue():*/
data.get(i).add(cell.getBooleanCellValue() + "");

/**И когда тип ячейки - ФОРМУЛА, мы можем использовать метод getCellFormula():*/
data.get(i).add(cell.getCellFormula() + "");

    /**Запись в Excel*/

}
