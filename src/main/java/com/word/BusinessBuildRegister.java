package com.word;

import lombok.Getter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.util.ArrayList;
import java.util.List;

/**
 * 业务注册
 * @author liuyongjun
 * @version 1.0
 * @Description: TODO
 * @date 2019/9/17 17:21
 */
@Getter
public class BusinessBuildRegister {


    private List<Business> businessList = new ArrayList<>();

    public void register(Business ... businessArr) throws WordException, IllegalAccessException, InstantiationException {
        for (Business business: businessArr) {
            businessList.add(business);
        }
    }

    public void replaceWord(XWPFDocument document){
        for (Business business:businessList) {
            ParagraphBuild build = business.paragraphBuild(document);
            build.replaceParagraph(document);
        }
    }
}
