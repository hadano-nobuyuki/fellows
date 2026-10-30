package com.example.sample01.model;

/** 注文情報詳細画面の1行分(表示用に整形済み)。 */
public class Order {

    private final String orderDate;
    private final String orderNo;
    private final int branchNo;
    private final String quantity;
    private final String coupon;
    private final boolean couponMinus;

    public Order(String orderDate, String orderNo, int branchNo, String quantity,
                 String coupon, boolean couponMinus) {
        this.orderDate = orderDate;
        this.orderNo = orderNo;
        this.branchNo = branchNo;
        this.quantity = quantity;
        this.coupon = coupon;
        this.couponMinus = couponMinus;
    }

    public String getOrderDate() { return orderDate; }
    public String getOrderNo() { return orderNo; }
    public int getBranchNo() { return branchNo; }
    public String getQuantity() { return quantity; }
    /** 符号付き・カンマ区切り。NULL の場合は空文字。 */
    public String getCoupon() { return coupon; }
    public boolean isCouponMinus() { return couponMinus; }
}
