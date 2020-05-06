package com.kland.common.other;

import com.aspose.words.*;

import java.io.FileInputStream;

public class ReplaceAndInsertImage implements IReplacingCallback {
    private String url;

    public ReplaceAndInsertImage(String url) {
        this.url = url;
    }

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    @Override
    public int replacing(ReplacingArgs replacingArgs) throws Exception {
        Node node = replacingArgs.getMatchNode();
        Document document = (Document) node.getDocument().getDocument();
        DocumentBuilder builder = new DocumentBuilder(document);
        // 将光标移动到指定节点
        builder.moveTo(node);
        builder.insertImage(new FileInputStream(url));
        return ReplaceAction.REPLACE;
    }
}
