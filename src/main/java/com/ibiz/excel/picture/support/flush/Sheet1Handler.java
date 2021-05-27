package com.ibiz.excel.picture.support.flush;

import com.ibiz.excel.picture.support.model.Cell;
import com.ibiz.excel.picture.support.model.MergeCell;
import com.ibiz.excel.picture.support.model.Row;
import com.ibiz.excel.picture.support.model.Sheet;
import com.ibiz.excel.picture.support.util.StringUtils;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Set;

/**
 * 图片
 * @auther 喻场
 * @date 2020/7/618:33
 */
public class Sheet1Handler implements InvocationHandler {
    private IRepository target;
    public Sheet1Handler(IRepository proxy) {
        this.target = proxy;
    }
    @Override
    public Object invoke(Object proxy, Method method, Object[] args) throws Throwable {
        Sheet sheet = (Sheet)args[0];
        if (method.getName().equals("write")) {
            List<Row> rows = sheet.getRows();
            if (!rows.isEmpty()) {
                //未刷新过说明没有写入过流,这里主要为了写表头
                //如果写过了,则从脚标1开始,原因是为了对比合并单元格在row1中保存上一次刷新的最后一条数据
                int subIndex = !sheet.hasFlush() ? 0 : 1;
                setMergeCell(sheet, rows);
                rows.subList(subIndex, rows.size()).stream().forEach(r -> writeSheetXML(r));
            }
        } else if (method.getName().equals("close")) {
            setEndSheetData();
            setMergeContent(sheet);
        }
        return method.invoke(target, args);
    }

    /**
     * 增加合并列内容
     * @param sheet
     */
    private void setMergeContent(Sheet sheet) {
        Set<String> colCells = sheet.getColCells();
        StringBuilder content = new StringBuilder();
        if (sheet.isAutoMergeCell() && !sheet.getMergeCells().isEmpty()) {
            content.append("<mergeCells count=\"").append(sheet.getMergeCells().size()).append("\">");
            sheet.getMergeCells().stream().forEach(m -> {
                colCells.forEach(c -> {
                    content.append("<mergeCell ref=\"").append(c).append(m.getStartRowNumber()).append(":")
                            .append(c).append(m.getEndRowNumber()).append("\"/>");
                });
            });
            content.append("</mergeCells>");
        }
        target.append(content.toString());
    }

    /**
     * 列结尾
     */
    private void setEndSheetData() {
        target.append("</sheetData>");
    }

    private void setMergeCell(Sheet sheet, List<Row> rows) {
        String oldValue = "";
        for(int i = 0; i < rows.size(); i++){
            Row row = rows.get(i);
            if(null == row){
                continue;
            }
            List<Cell> cells = row.getCells();
            int mergeCellNumber = sheet.getMergeCellNumber();
            if(mergeCellNumber >= cells.size()){
                //合并列超出范围,不进行合并
                break;
            }
            Cell cell = cells.get(mergeCellNumber);
            if((StringUtils.isNotBlank(oldValue) && oldValue.equals(cell.getValue()))){
                MergeCell mergeCell = null;
                if(!sheet.getMergeCells().isEmpty()){
                    mergeCell = sheet.getMergeCells().getLast();
                }
                if (null == mergeCell) {
                    mergeCell = new MergeCell();
                }
                int endRowNumber = mergeCell.getEndRowNumber();
                if (row.getRowNumber() == endRowNumber) {
                    //与上一个合并对象合并
                    mergeCell.endRowNumber(++endRowNumber);
                } else {
                    mergeCell = new MergeCell().startRowNumber(row.getRowNumber()).endRowNumber(row.getRowNumber() + 1);
                    sheet.getMergeCells().add(mergeCell);
                }
            }
            oldValue = StringUtils.isBlank(cell.getValue()) ? "" : cell.getValue();
        }
    }

    private void writeSheetXML(Row row) {
        StringBuilder content = new StringBuilder();
        //customHeight=1 使用自定义高度
        content.append("<row r=\"").append(row.getRowNumber() + 1).append("\" ht=\"").append(row.getHeight())
                .append("\" customHeight=\"1\"").append(" spans=\"1:").append(row.getCells().size()).append("\">");
        row.getCells().forEach(c ->
            content.append("<c r=\"").append(c.getColNumber()).append("\" s=\"1\" t=\"s\">")
                    .append("<v>").append(c.getColDataNumber()).append("</v></c>")
        );
        content.append("</row>");
        target.append(content.toString());
    }

}
