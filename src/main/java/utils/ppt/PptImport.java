package utils.ppt;

import com.spire.presentation.*;
import com.spire.presentation.collections.TextAnimationCollection;
import com.spire.presentation.drawing.animation.AnimationEffect;
import com.spire.presentation.drawing.animation.AnimationEffectType;
import com.spire.presentation.drawing.animation.ParagraphBuildType;
import org.apache.poi.hslf.usermodel.*;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static utils.ppt.DataProcessing.*;
import static utils.ppt.DataProcessingX.*;

/**
 * @Author: zdl
 * @Date: 2021/9/6 15:28
 */
public class PptImport {
    final static String url = "data/test.pptx";
    //获取动画效果
    private static void getAnimation(Map<String,String> animationMap,Map<String,String> mp4Map) throws Exception {
        final Map<Long,String> idText = new HashMap<Long,String>();
        Presentation presentation = new Presentation();
        presentation.loadFromFile(url);
        for (int c = 0; c < presentation.getSlides().getCount(); c++) {
            ISlide slide = presentation.getSlides().get(c);
            for(int i = 0; i< slide.getShapes().getCount(); i++) {
                IShape shape = slide.getShapes().get(i);
                if ((shape instanceof IVideo)) {
                    IVideo video = (IVideo) shape;
                    try {
                        File folder = new File("data");
                        video.getEmbeddedVideoData().saveToFile("data/"+video.getName()+ ".mp4");
                        mp4Map.put(video.getName(),"data/"+video.getName()+ ".mp4");
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
    private static void getElement(Map<String,String> animationMap,Map<String,String> mp4Map) throws IOException {
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
                List<XSLFShape> shapes = slide.getShapes();
                // 循环页内所有元素
                for (int f = 0; f < shapes.size(); f++) {
                    XSLFShape shape = shapes.get(f);
                    processingDataX(shape,pageWidthProportion,pageHeightProportion,f*i,animationMap,mp4Map);
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
            ,Map<String,String> animationMap,Map<String,String> mp4Map) {
        if (shape instanceof XSLFTextBox) {
            // 文字
            textProcessX(shape,pageWidthProportion,pageHeightProportion,count,animationMap);
        } else if (shape instanceof XSLFAutoShape) {
            // 图形
            autoShapeProcessX(shape,pageWidthProportion,pageHeightProportion,count,animationMap);
        } else if (shape instanceof XSLFPictureShape) {
            // 图片
            pictureProcessX(shape,pageWidthProportion,pageHeightProportion,count,animationMap,mp4Map);
        } else if (shape instanceof XSLFGroupShape) {
            System.out.println("--组合图形--");
            XSLFGroupShape groupShape = (XSLFGroupShape) shape;
            for(XSLFShape xslfShape:groupShape.getShapes()){
                if(xslfShape != null){
                    processingDataX(xslfShape,pageWidthProportion,pageHeightProportion,count,animationMap,mp4Map);
                }
            }
        }
    }

    public static void main(String[] args) throws Exception {
        Map<String,String> animationMap = new HashMap<String, String>();
        Map<String,String> mp4Map = new HashMap<String, String>();
        getAnimation(animationMap,mp4Map);
        getElement(animationMap,mp4Map);
    }
}
