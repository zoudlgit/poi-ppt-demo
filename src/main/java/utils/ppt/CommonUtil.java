package utils.ppt;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.sl.usermodel.PaintStyle;

/**
 * @Author: zdl
 * @Date: 2021/9/6 15:28
 */
public class CommonUtil {
    public static String getColor(String[] strings){
        String color = "rgb(";
        color = color + strings[0].substring(strings[0].lastIndexOf("=")+1) + "," +
                strings[1].substring(strings[1].lastIndexOf("=")+1) + "," +
                strings[2].substring(strings[2].lastIndexOf("=")+1,strings[2].length()-1) +")";
        return color;
    }
    public static String getColor1(PaintStyle paintStyle){
        Object object = JSON.parse(JSON.toJSONString(paintStyle));
        Object jsonObject = ((JSONObject) object).get("solidColor");
        Object colorObject = ((JSONObject) jsonObject).get("color");
        String color = "rgb(";
        color = color + ((JSONObject) colorObject).get("r").toString() + "," + ((JSONObject) colorObject).get("g").toString()+ ","
                + ((JSONObject) colorObject).get("b").toString() +")";
        return color;
    }
}
