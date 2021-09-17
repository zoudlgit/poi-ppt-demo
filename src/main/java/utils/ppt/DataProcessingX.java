package utils.ppt;

import com.alibaba.fastjson.JSON;
import org.apache.poi.xslf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static utils.ppt.CommonUtil.getColor;
import static utils.ppt.CommonUtil.getColor1;

/**
 * @Author: zdl
 * @Date: 2021/9/17 9:49
 */
public class DataProcessingX {
    public static void autoShapeProcessX(XSLFShape shape, double pageWidthProportion, double pageHeightProportion, int i, Map<String,String> animationMap) {
        System.out.println("-------------图形处理-------------");
        String graphicType = getGraphicType(shape);
        XSLFAutoShape autoShape = (XSLFAutoShape) shape;
        Map<String,String> styleMap = new HashMap();
        //圆
        if("CIRCLE".equals(graphicType)){
            styleMap.put("width",autoShape.getAnchor().getWidth()*2 +"px");
            styleMap.put("height",autoShape.getAnchor().getHeight()*2 +"px");
            styleMap.put("left",autoShape.getAnchor().getMinX()/pageWidthProportion-20 +"px");
            styleMap.put("top",autoShape.getAnchor().getMinY()/pageHeightProportion-20 +"px");
            styleMap.put("cx",autoShape.getAnchor().getWidth() +"px");
            styleMap.put("cy",autoShape.getAnchor().getHeight() +"px");
            styleMap.put("rx",autoShape.getAnchor().getWidth() +"px");
            styleMap.put("ry",autoShape.getAnchor().getHeight() +"px");
        } else {
            styleMap.put("width",autoShape.getAnchor().getWidth()/pageWidthProportion +"px");
            styleMap.put("height",autoShape.getAnchor().getHeight()/pageHeightProportion +"px");
            styleMap.put("left",autoShape.getAnchor().getMinX()/pageWidthProportion +"px");
            styleMap.put("top",autoShape.getAnchor().getMinY()/pageHeightProportion +"px");
        }
        // 空白图形 标题一不处理返回
        if(autoShape.getFillColor() != null){
            styleMap.put("fill",getColor(autoShape.getFillColor().toString().split(",")));
        } else {
            return;
        }
        styleMap.put("z-index",i+"");
        styleMap.put("strokeWidth",autoShape.getLineWidth() +"px");
        if(autoShape.getLineColor() != null){
            styleMap.put("stroke", getColor(autoShape.getLineColor().toString().split(",")));
        }
        styleMap.put("strokeDasharray",autoShape.getStrokeStyle().getLineDash().toString());
        if(autoShape.getFillColor() != null && autoShape.getFillColor().getTransparency() == 3){
            styleMap.put("opacity",20*autoShape.getFillColor().getAlpha()/51+"");
        } else {
            styleMap.put("opacity",100+"");
        }
        System.out.println("图形类型："+graphicType);
        System.out.println("图形样式："+JSON.toJSONString(styleMap));
        System.out.println("图形动画效果："+animationMap.get(autoShape.getShapeName()));
    }
    public static void textProcessX(XSLFShape shape,double pageWidthProportion,double pageHeightProportion,int i,Map<String,String> animationMap) {
        System.out.println("-------------文字处理-------------");
        XSLFTextBox textBox = (XSLFTextBox) shape;
        List<XSLFTextParagraph> XSLFTextParagraphs = textBox.getTextParagraphs();
        Map<String,String> styleMap = new HashMap();
        XSLFTextRun textRuns = XSLFTextParagraphs.get(0).getTextRuns().get(0);
        styleMap.put("width",textBox.getAnchor().getWidth()/pageWidthProportion +"px");
        styleMap.put("height",textBox.getAnchor().getHeight()/pageHeightProportion +"px");
        styleMap.put("left",textBox.getAnchor().getMinX()/pageWidthProportion +"px");
        styleMap.put("top",textBox.getAnchor().getMinY()/pageHeightProportion +"px");
        if(textRuns.getFontColor() != null){
            styleMap.put("color",getColor1(textRuns.getFontColor()));
        }
        styleMap.put("z-index",i+"");
        styleMap.put("border-width",textBox.getLineWidth() +"px");
        if(textBox.getFillColor() != null){
            styleMap.put("background-color",getColor(textBox.getFillColor().toString().split(",")));
        }
        if(textBox.getLineColor() != null){
            styleMap.put("border-color",getColor(textBox.getLineColor().toString().split(",")));
        }
        String content = "";
        String style = "\"";
        if(textRuns.getFontSize() != null){
            style = style+"font-size:" + textRuns.getFontSize()+"px;";
        }
        if(textRuns.getFontColor() != null){
            style = style+"color:" + getColor1(textRuns.getFontColor())+";";
        }
        if(textRuns.getFontFamily() != null){
            style = style+"font-family:" + textRuns.getFontFamily()+";";
        }
        for(String string:textBox.getText().split("\n")){
            content = content + "<div><span style="+ style +"\">" +string+"</span></div>";
        }
        System.out.println("文字内容：" + textBox.getText());
        System.out.println("文字外部样式：" + JSON.toJSONString(styleMap));
        System.out.println("文字内部样式：" + content);
        System.out.println("文字动画效果："+animationMap.get(textBox.getShapeName()));
    }
    public static void pictureProcessX(XSLFShape shape,double pageWidthProportion,double pageHeightProportion,int i,
                                       Map<String,String> animationMap,Map<String,String> mp4Map) {
        XSLFPictureShape pictureShape = (XSLFPictureShape) shape;
        if(mp4Map.get(pictureShape.getShapeName()) != null){
            System.out.println("-------------视频处理-------------");
            Map<String,String> pictureMap = new HashMap();
            pictureMap.put("width",pictureShape.getAnchor().getWidth()/pageWidthProportion +"px");
            pictureMap.put("height",pictureShape.getAnchor().getHeight()/pageHeightProportion +"px");
            pictureMap.put("left",pictureShape.getAnchor().getMinX()/pageWidthProportion +"px");
            pictureMap.put("top",pictureShape.getAnchor().getMinY()/pageHeightProportion +"px");
            pictureMap.put("rorateX",(pictureShape.getAnchor().getX())  + "px");
            pictureMap.put("rorateY",(pictureShape.getAnchor().getY()) + "px");
            String css =JSON.toJSONString(pictureMap);
            System.out.println("视频样式："+css);
            System.out.println("视频地址："+"data/"+pictureShape.getShapeName()+".mp4");
            System.out.println("视频动画效果："+animationMap.get(pictureShape.getShapeName()));
        } else {
            System.out.println("-------------图片处理-------------");
            ByteArrayInputStream bais = new ByteArrayInputStream(pictureShape.getPictureData().getData());
            BufferedImage bi1 = null;
            try {
                bi1 = ImageIO.read(bais);
                File w2 = new File("data/"+pictureShape.getShapeName()+".png");//可以是jpg,png,gif格式
                ImageIO.write(bi1, "png", w2);//不管输出什么格式图片，此处不需改动
            } catch (IOException e) {
                e.printStackTrace();
                System.out.println("path");
            }
            Map<String,String> pictureMap = new HashMap();
            pictureMap.put("width",pictureShape.getAnchor().getWidth()/pageWidthProportion +"px");
            pictureMap.put("height",pictureShape.getAnchor().getHeight()/pageHeightProportion +"px");
            pictureMap.put("left",pictureShape.getAnchor().getMinX()/pageWidthProportion +"px");
            pictureMap.put("top",pictureShape.getAnchor().getMinY()/pageHeightProportion +"px");
            pictureMap.put("rorateX",(pictureShape.getAnchor().getX())  + "px");
            pictureMap.put("rorateY",(pictureShape.getAnchor().getY()) + "px");
            String css =JSON.toJSONString(pictureMap);
            System.out.println("图片样式："+css);
            System.out.println("图片地址："+"data/"+pictureShape.getShapeName()+".png");
            System.out.println("图片动画效果："+animationMap.get(pictureShape.getShapeName()));
        }
    }

    private static String getGraphicType(XSLFShape shape) {
        if("RECT".equals(shape.getShapeName())){
            return "SQURE";
        } else if("ELLIPSE".equals(shape.getShapeName())){
            return "CIRCLE";
        } else if("TRIANGLE".equals(shape.getShapeName())){
            return "TRIANGLE";
        }
        return null;
    }
}
