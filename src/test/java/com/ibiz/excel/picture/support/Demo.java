package com.ibiz.excel.picture.support;

import com.ibiz.excel.picture.support.model.Sheet;
import com.ibiz.excel.picture.support.model.Workbook;
import java.io.*;
import java.util.Date;
import java.util.UUID;
import java.util.stream.IntStream;

/**
 * @auther 喻场
 * @date 2020/7/817:29
 */
public class Demo {
    static final String CURRENT_PATH = "/Users/ronaldo/Desktop/";//Demo.class.getClassLoader().getResource(".").getPath();
    static final String picture1 = "http://h0.beicdn.com/open202121/83c9da8eca014eba_180x179.png";
    static final String picture2 = "http://h0.beicdn.com/open202121/915392253f5e5695_180x177.png";
    static final String picture3 = "http://h0.beicdn.com/open202121/b5f7fe014ac9e1f3_180x179.png";
    static final String picture4 = "http://h0.beicdn.com/open202121/70a2c51d55ccc213_183x176.png";
    static final String picture5 = "http://h0.beicdn.com/open202121/1fa74535583a18ad_183x177.png";
    public static void main(String[] args) throws IOException {
        Date start = new Date();
        Workbook workBook = Workbook.getInstance(200);
        Sheet sheet = workBook.createSheet("导出个表");
        IntStream.range(1, 15).forEach(r -> {
            String pre = "a";
            if (r % 3 == 0) {
                pre = "b";
            }
            User u1 = new User(pre + "1", 12, pre + "1", picture1);
            u1.setDepartment1(pre + "1");
            u1.setDepartment2(pre + "1");
            u1.setDepartment3(pre + "1");
            u1.setDepartment4(pre + "1");
            u1.setDepartment5(pre + "1");
            u1.setDepartment6(pre + "1");
            u1.setDepartment7(pre + "1");
            u1.setPicture1(picture2);
            u1.setPicture2(picture3);
            u1.setPicture3(picture4);
            if (r % 3 == 0 || r == 2) {
                u1.setPicture(null);
                u1.setPicture1(null);
                u1.setPicture2(null);
                u1.setPicture3(null);
            }
            sheet.createRow(u1);
        });
        File file = createFile();
        OutputStream os = new FileOutputStream(file);
        workBook.write(os);
        workBook.close();
        Date end = new Date();
        System.out.println("file capital :" + (file.length()/1024/1024) +"M  name :" + file.getName());
        System.out.println("file cost time :" + (end.getTime() - start.getTime()));
        os.close();
    }

    private static File createFile() {
        String dir = CURRENT_PATH + "excel/";
        File file = new File(dir);
        if (!file.exists()) {
            file.mkdir();
        }
        String name = CURRENT_PATH + "excel/" + UUID.randomUUID() + ".xlsx";
        return new File(name);
    }

}
