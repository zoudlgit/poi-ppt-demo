# poi-ppt-demo＜/br＞
最近有一个导入ppt,识别ppt内所有元素需求，翻了一些资料都没有特别好的demo,官方文档很官方...,所以打算自己写一个＜/br＞
主要用的是poi 4.1,因为项目里之前引入的就是4.1所以就不用最新的了＜/br＞
官网：http://poi.apache.org/apidocs/4.1/＜/br＞
还有Spire.Presentation for Java 付费插件用到了保存视频和读取动画效果两个功能＜/br＞
官网：https://www.e-iceblue.cn/spirepresentationforjava/spire-presentation-for-java-program-guide-content-html.html＜/br＞

记录踩坑＜/br＞
1.maven下载spire.presentation包不稳定，我把它放到了私有依赖库＜/br＞
2.因为ppt的页面大小和网页的大小可能不一样,会导致位置不对,所以要先获取比例＜/br＞
3.尽量输出png格式图片，镂空图片还没解决＜/br＞
4.ppt和pptx是两套处理逻辑＜/br＞
5.文字多层样式多次处理＜/br＞
6.艺术字默认转换为图片＜/br＞
7.ppt格式不支持视频、音频＜/br＞
8.poi 没有视频对象  我是没找到。。 用spire.presentation获取的视频＜/br＞