/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelread;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import static org.junit.Assert.assertEquals;
import org.junit.Test;

/**
 * 读取
 *
 * @author ZhangLiKun
 * @mail likun_zhang@yeah.net
 * @date 2013-5-8
 */
public class ExcelReaderTest {

  @Test
  public void testRead() throws FileNotFoundException, IOException {
    Workbook wb = ExcelReader.createWb("test2.xls");

    // 获取Workbook中Sheet个数
    int sheetTotal = wb.getNumberOfSheets();
    Debug.printf("工作簿中的工作表个数为：{}", sheetTotal);

    for (int k = 0; k < sheetTotal; k++) {
      Sheet sheet = ExcelReader.getSheet(wb, k);
      List<Object[]> list = ExcelReader.listFromSheet(sheet);
      Debug.printf(list, new ToString<Object[]>() {
        private int index = 1;

        @Override
        public String toString(Object[] t) {

          if (t == null || t.length == 0) {
            return StringUtils.EMPTY;
          }
          StringBuffer sb = new StringBuffer(index++ + ":");
          for (int i = 0, len = t.length; i < len; i++) {
            sb.append(t[i] + ",");
          }
          return sb.toString();
        }
      });
    }
  }

}
