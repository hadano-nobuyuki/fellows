package com.example.sample01.model;

/** 商品一覧画面の1行分(表示用に整形済み)。 */
public class Product {

    private final String code;
    private final String name;
    private final String unitPrice;
    private final String category;
    private final String relatedCode;
    private final int orderCount;

    public Product(String code, String name, String unitPrice, String category,
                   String relatedCode, int orderCount) {
        this.code = code;
        this.name = name;
        this.unitPrice = unitPrice;
        this.category = category;
        this.relatedCode = relatedCode;
        this.orderCount = orderCount;
    }

    public String getCode() { return code; }
    public String getName() { return name; }
    public String getUnitPrice() { return unitPrice; }
    public String getCategory() { return category; }
    /** 関連商品コード(無い場合は null)。 */
    public String getRelatedCode() { return relatedCode; }
    public int getOrderCount() { return orderCount; }
}
