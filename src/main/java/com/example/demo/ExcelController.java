package com.example.demo;

import com.sun.org.apache.regexp.internal.RE;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/**
 * <p>
 * 最近有一个需求是需要从n张excel导出表数据，生成sql语句。简单记录下实现过程。
 * </p>
 *
 * @author wangdejian
 * @since 2018/3/9
 */
@RestController
public class ExcelController {

    @RequestMapping(value = "poiExportExcel")
    public void poiExportExcel() throws IOException {
        Path dir = Paths.get("C:\\gitworkspace\\excel-es-mysql\\src\\main\\resources\\excel");

        // 获取到目录下所有的excel.
        Files.walkFileTree(dir, new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs)
                    throws IOException {
                if (file.toString().endsWith(".xls")) {
                    // 获取到当前excel名称:即个体工商户.xls
                    String excelName = file.getFileName().toString();
                    writeSqlStatement(excelName.substring(0, excelName.length() - 4));
                }
                return super.visitFile(file, attrs);
            }
        });
    }

    private void writeSqlStatement(String tableName) throws IOException {
        List<String> list = new ArrayList<>();

        //Excel文件
        HSSFWorkbook book = new HSSFWorkbook(new FileInputStream(
                ResourceUtils.getFile("classpath:excel/" + tableName + ".xls")));

        HSSFSheet sheet = book.getSheetAt(0);
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 1; i < lastRowNum + 1; i++) {
            HSSFRow row = sheet.getRow(i);
            // 如果excel没有清除格式，那么无法获取到真正的列长度。导致空指针异常；判断如果为空说明已经到最后.
            Optional<HSSFCell> cell = Optional.ofNullable(row.getCell(1));
            if (!cell.isPresent()) {
                break;
            }
            String name = row.getCell(1).getStringCellValue(); //url

            String insertSql = "INSERT INTO `t_field` VALUES ('" + tableName + "', '" + tableName + "', '" + name + "', '" + name + "');";
            list.add(insertSql);
        }
        // 写整理好的数据到templates目录下
        File file = new File("C:\\gitworkspace\\excel-es-mysql\\src\\main\\resources\\templates\\" + tableName + ".txt");
        if (!file.exists()) {
            file.createNewFile();
        }

        // 写数据到txt中
        writeInsertToTxt(file, list);
    }

    private void writeInsertToTxt(File file, List<String> list) throws IOException {
        FileOutputStream fos = new FileOutputStream(file);
        BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));
        for (String s : list) {
            bw.write(s);
            bw.newLine();
        }
        bw.close();
    }

}
