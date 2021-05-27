package com.ibiz.excel.picture.support.flush;

import com.ibiz.excel.picture.support.util.CovertUtil;
import com.ibiz.excel.picture.support.constants.Alias;
import com.ibiz.excel.picture.support.constants.WorkbookConstant;
import com.ibiz.excel.picture.support.model.Picture;
import com.ibiz.excel.picture.support.model.Sheet;
import com.ibiz.excel.picture.support.module.RelationShip;
import com.ibiz.excel.picture.support.util.FileUtil;
import com.ibiz.excel.picture.support.util.MD5Digester;
import com.ibiz.excel.picture.support.util.StringUtils;
import org.apache.tools.ant.util.FileUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.net.HttpURLConnection;
import java.net.URI;
import java.net.URL;
import java.util.*;

/**
 * 图片
 * @auther 喻场
 * @date 2020/7/618:33
 */
public class DrawingXmlRelsHandler implements InvocationHandler {
    private Logger logger = LoggerFactory.getLogger(this.getClass());
    private IRepository target;
    public DrawingXmlRelsHandler(IRepository proxy) {
        this.target = proxy;
    }
    @Override
    public Object invoke(Object proxy, Method method, Object[] args) throws Throwable {
        if (method.getName().equals("write")) {
            Sheet sheet = (Sheet)args[0];
            List<Picture> pictures = sheet.getPictures();
            pictures.stream().filter(Objects::nonNull).forEach(p -> {
                if (StringUtils.isNotBlank(p.getPicturePath())) {
                    copyPictureAppendDrawingRelsXML(sheet, p);
                }
            });
        }
        return method.invoke(target, args);
    }

    /**
     * 把图片copy到media目录下
     * @param sheet
     * @param picture
     */
    private void copyPictureAppendDrawingRelsXML(Sheet sheet, Picture picture) {
        try {
            String fileName = "/tmp" + picture.getPicturePath().substring(picture.getPicturePath().lastIndexOf("/"));
            saveUrlAs(picture.getPicturePath(), fileName);

            File srcPicture = new File(fileName);
            String md5 = MD5Digester.digestMD5(srcPicture);
            Integer drawingSequence = sheet.getWorkbook().getImageCache().get(md5);
            if (Objects.nonNull(drawingSequence)) {
                RelationShip relationShip = new RelationShip("rId" + drawingSequence, WorkbookConstant.MEDIA_IMAGE_TYPE, "../media/image" + drawingSequence +".png");
                picture.setRembed(drawingSequence);
                target.append(CovertUtil.covert(relationShip));
                return;
            }
            File media = sheet.getSheetContext().getRepositoryHolder().get(Alias.MEDIA).getFile();
            drawingSequence = sheet.getDrawingSequence();
            File destPicture = new File(media, "image"+ drawingSequence + ".png");
            FileUtil.copyFile(srcPicture, destPicture);
            picture.setRembed(drawingSequence);
            RelationShip relationShip = new RelationShip("rId" + drawingSequence, WorkbookConstant.MEDIA_IMAGE_TYPE, "../media/image" + drawingSequence +".png");
            //先把drawingSequence放入缓存,因为从缓存中获取时设置drawingSequence再+1对应图片。实际图片的drawingSequence不变
            sheet.getWorkbook().getImageCache().put(md5, drawingSequence);
            drawingSequence++;
            sheet.setDrawingSequence(drawingSequence);
            target.append(CovertUtil.covert(relationShip));
        } catch (IOException e) {
            logger.error("图片copy到media目录下异常",e);
        } catch (Exception e) {
            logger.error("图片copy到media目录下异常",e);
        }
    }

    public boolean saveUrlAs(String photoUrl, String fileName) {
        try {
            URL url = new URL(photoUrl);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            DataInputStream in = new DataInputStream(connection.getInputStream());
            DataOutputStream out = new DataOutputStream(new FileOutputStream(fileName));
            byte[] buffer = new byte[4096];
            int count = 0;
            while ((count = in.read(buffer)) > 0) {
                out.write(buffer, 0, count);
            }
            out.close();
            in.close();
            return true;
        }
        catch (Exception e) {
            return false;
        }
    }


}
