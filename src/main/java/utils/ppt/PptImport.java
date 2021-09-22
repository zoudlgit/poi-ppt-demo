package utils.ppt;

import com.spire.presentation.*;
import com.spire.presentation.collections.TextAnimationCollection;
import com.spire.presentation.drawing.animation.AnimationEffect;
import com.spire.presentation.drawing.animation.AnimationEffectType;
import com.spire.presentation.drawing.animation.ParagraphBuildType;
import org.apache.poi.hslf.usermodel.*;
import org.apache.poi.xslf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static utils.ppt.CommonUtil.getColor;
import static utils.ppt.DataProcessing.*;
import static utils.ppt.DataProcessingX.*;

/**
 * @Author: zdl
 * @Date: 2021/9/6 15:28
 */
public class PptImport {
    final static String url = "data/test.pptx";
    //获取动画效果
    private static void getAnimation(Map<String,String> animationMap,Map<String,String> mp4Map,Map<String,String> mp3Map
            ,Map<Integer,byte[]> bgImageMap) throws Exception {
        final Map<Long,String> idText = new HashMap<Long,String>();
        Presentation presentation = new Presentation();
        presentation.loadFromFile(url);
        for (int c = 0; c < presentation.getSlides().getCount(); c++) {
            ISlide slide = presentation.getSlides().get(c);
            if(slide.getSlideBackground().getFill().getPictureFill().getPicture().getEmbedImage() != null){
                bgImageMap.put(c,slide.getSlideBackground().getFill().getPictureFill().getPicture().getEmbedImage().getData());
            }
            for(int i = 0; i< slide.getShapes().getCount(); i++) {
                IShape shape = slide.getShapes().get(i);
                if ((shape instanceof IVideo)) {
                    IVideo video = (IVideo) shape;
                    try {
                        video.getEmbeddedVideoData().saveToFile("data/"+video.getName()+ ".mp4");
                        mp4Map.put(video.getName(),"data/"+video.getName()+ ".mp4");
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                } else if ((shape instanceof IAudio)) {
                    IAudio audio = (IAudio) shape;
                    try {
                        audio.getData().saveToFile("data/"+audio.getName()+ ".mp3");
                        mp3Map.put(audio.getName(),"data/"+audio.getName()+ ".mp3");
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }
            //记录元素名map
            for (Object shape : (slide).getShapes()) {
                if (shape instanceof IAutoShape) {
                    idText.put(((IAutoShape) shape).getId(),((IAutoShape) shape).getTextFrame().getParagraphs().get(0).getText());
                    if(ShapeType.RECTANGLE.equals(((IAutoShape) shape).getShapeType())){
                        String text = "";
                        for(int i=0;i<((IAutoShape) shape).getTextFrame().getParagraphs().size();i++){
                            text = text+((IAutoShape) shape).getTextFrame().getParagraphs().get(i).getText();
                        }
                        idText.put(((IAutoShape) shape).getId(),((IAutoShape) shape).getName());
                        animationMap.put(text,"");
                    } else {
                        idText.put(((IShape) shape).getId(),((IShape) shape).getName());
                        animationMap.put(((IShape) shape).getName(),"");
                    }
                } else if (shape instanceof SlidePicture) {
                    idText.put(((SlidePicture) shape).getId(),((SlidePicture) shape).getName());
                    animationMap.put(((SlidePicture) shape).getName(),"");
                } else if ((shape instanceof IVideo)) {
                    idText.put(((IVideo) shape).getId(),((IVideo) shape).getName());
                    animationMap.put(((IVideo) shape).getName(),"");
                } else if (shape instanceof GroupShape){
                    idText.put(((GroupShape) shape).getId(),((GroupShape) shape).getName());
                    animationMap.put(((GroupShape) shape).getName(),"");
                }
            }
            // slide.getTimeline()是所有的动画效果
            for (int i = 0; i < slide.getTimeline().getMainSequence().getCount(); i++) {
                AnimationEffect animationEffect = slide.getTimeline().getMainSequence().get(i);
                //预设类型，比如Entrance,Emphasis,Exit,Path
                String animation = "";
                String presetClassType = animationEffect.getPresetClassType().getName();
                animation = animation+"入口:"+presetClassType+";";
                //获取动画效果类型
                AnimationEffectType animationEffectType= animationEffect.getAnimationEffectType();
                animation = animation+"动画效果类型:"+animationEffectType+";";
                //获取目标Shape
                com.spire.presentation.Shape shape = animationEffect.getShapeTarget();
                //获取动画效果子类型
                String subType = animationEffect.getSubtype().getName();
                animation = animation+"方式:"+subType+";";
                //获取Color
                Color color = animationEffect.getColor();
                animation = animation+"color:"+color+";";
                //当动画效果类型为Faded_Zoom时，获取vanishing point（消失点）
                if (animationEffectType.equals(AnimationEffectType.FADED_ZOOM)) {
                    String vanishingPointName = animationEffect.getVanishingPoint().getName();
                    System.out.println("vanishingPointName:"+vanishingPointName);
                    animation = animation+"color:"+color+";";
                }
                //获取WAVE动画效果
                if (animationEffectType.equals(AnimationEffectType.WAVE)) {
                    TextAnimationCollection textAnimations = slide.getTimeline().getTextAnimations();
                    if (textAnimations.size() > 0) {
                        for (int j = 0; j < textAnimations.size(); j++) {
                            ParagraphBuildType buildType = textAnimations.get(j).getParagraphBuildType();
                        }
                    }
                }
                animationMap.put(idText.get(shape.getId()),animation);
            }
        }
    }
    private static void getElement(Map<String,String> animationMap,Map<String,String> mp4Map,Map<String,String> mp3Map,
                                   Map<Integer,byte[]> bgImageMap) throws IOException {
        FileInputStream fis = new FileInputStream(url);
        if("ppt".equals(url.substring(url.lastIndexOf(".")+1))){
            //实例化ppt
            HSLFSlideShow ppt = new HSLFSlideShow(fis);
            //获取比例
            double pageWidthProportion = ppt.getPageSize().getWidth()/960;
            double pageHeightProportion = ppt.getPageSize().getHeight()/540;
            //循环每页
            for (int i = 0; i < ppt.getSlides().size(); i++) {
                System.out.println("-----第"+i+"页-----");
                HSLFSlide slide = ppt.getSlides().get(i);
                //fillType: 0.颜色, 1.图片
                if(slide.getBackground().getFill().getFillType() == 3){
                    ByteArrayInputStream bais = new ByteArrayInputStream(slide.getBackground().getFill().getPictureData().getData());
                    BufferedImage bi1 = null;
                    try {
                        bi1 = ImageIO.read(bais);
                        File w2 = new File("data/background"+i+".png");//可以是jpg,png,gif格式
                        ImageIO.write(bi1, "png", w2);//不管输出什么格式图片，此处不需改动
                    } catch (IOException e) {
                        e.printStackTrace();
                        System.out.println("path");
                    }
                    System.out.println("背景图片");
                } else if (slide.getBackground().getFill().getFillType() == 0){
                    String[] strings = slide.getBackground().getFill().getForegroundColor().toString().split(",");
                    String color = getColor(strings);
                    System.out.println("背景颜色："+color);
                }
                List<HSLFShape> shapes = slide.getShapes();
                // 循环页内所有元素
                for (int f = 0; f < shapes.size(); f++) {
                    HSLFShape shape = shapes.get(f);
                    processingData(shape,pageWidthProportion,pageHeightProportion,f*i,animationMap);
                }
            }
        } else if("pptx".equals(url.substring(url.lastIndexOf(".")+1))){
            //实例化ppt
            XMLSlideShow ppt = new XMLSlideShow(fis);
            //获取比例
            double pageWidthProportion = ppt.getPageSize().getWidth()/960;
            double pageHeightProportion = ppt.getPageSize().getHeight()/540;
            //循环每页
            for (int i = 0; i < ppt.getSlides().size(); i++) {
                System.out.println("-----第"+i+"页-----");
                XSLFSlide slide = ppt.getSlides().get(i);
                //fillType: 0.颜色, 1.图片
                if(slide.getBackground().getFillColor() == null){
                    ByteArrayInputStream bais = new ByteArrayInputStream(bgImageMap.get(i));
                    BufferedImage bi1 = null;
                    try {
                        bi1 = ImageIO.read(bais);
                        File w2 = new File("data/background"+i+".png");//可以是jpg,png,gif格式
                        ImageIO.write(bi1, "png", w2);//不管输出什么格式图片，此处不需改动
                    } catch (IOException e) {
                        e.printStackTrace();
                        System.out.println("path");
                    }
                    System.out.println("背景图片");
                } else {
                    String[] strings = slide.getBackground().getFillColor().toString().split(",");
                    String color = getColor(strings);
                    System.out.println("背景颜色："+color);
                }
                List<XSLFShape> shapes = slide.getShapes();
                // 循环页内所有元素
                for (int f = 0; f < shapes.size(); f++) {
                    XSLFShape shape = shapes.get(f);
                    processingDataX(shape,pageWidthProportion,pageHeightProportion,f*i,animationMap,mp4Map,mp3Map);
                }
            }
        } else {
            System.out.println("请使用ppt、pptx格式文件");
        }
    }

    private static void processingData(HSLFShape shape,double pageWidthProportion,double pageHeightProportion,int count,Map<String,String> animationMap) {
        if (shape instanceof HSLFAutoShape) {
            // 图形
            autoShapeProcess(shape,pageWidthProportion,pageHeightProportion,count,animationMap);
        } else if (shape instanceof HSLFTextBox) {
            // 文字
            textProcess(shape,pageWidthProportion,pageHeightProportion,count,animationMap);
        } else if (shape instanceof HSLFPictureShape) {
            // 图片
            pictureProcess(shape,pageWidthProportion,pageHeightProportion,count,animationMap);
        } else if (shape instanceof HSLFGroupShape) {
            System.out.println("--组合图形--");
            HSLFGroupShape groupShape = (HSLFGroupShape) shape;
            for(HSLFShape hslfShape:groupShape.getShapes()){
                if(hslfShape != null){
                    processingData(hslfShape,pageWidthProportion,pageHeightProportion,count,animationMap);
                }
            }
        }
    }
    private static void processingDataX(XSLFShape shape,double pageWidthProportion,double pageHeightProportion,int count
            ,Map<String,String> animationMap,Map<String,String> mp4Map,Map<String,String> mp3Map) {
        if (shape instanceof XSLFTextBox) {
            // 文字
            textProcessX(shape,pageWidthProportion,pageHeightProportion,count,animationMap);
        } else if (shape instanceof XSLFAutoShape) {
            // 图形
            autoShapeProcessX(shape,pageWidthProportion,pageHeightProportion,count,animationMap);
        } else if (shape instanceof XSLFPictureShape) {
            // 图片
            pictureProcessX(shape,pageWidthProportion,pageHeightProportion,count,animationMap,mp4Map,mp3Map);
        } else if (shape instanceof XSLFGroupShape) {
            System.out.println("--组合图形--");
            XSLFGroupShape groupShape = (XSLFGroupShape) shape;
            for(XSLFShape xslfShape:groupShape.getShapes()){
                if(xslfShape != null){
                    processingDataX(xslfShape,pageWidthProportion,pageHeightProportion,count,animationMap,mp4Map,mp3Map);
                }
            }
        }
    }

    public static void main(String[] args) throws Exception {
        Map<String,String> animationMap = new HashMap<String, String>();
        Map<String,String> mp4Map = new HashMap<String, String>();
        Map<String,String> mp3Map = new HashMap<String, String>();
        Map<Integer,byte[]> bgImageMap = new HashMap<Integer, byte[]>();
        getAnimation(animationMap,mp4Map,mp3Map,bgImageMap);
        getElement(animationMap,mp4Map,mp3Map,bgImageMap);
    }
}
