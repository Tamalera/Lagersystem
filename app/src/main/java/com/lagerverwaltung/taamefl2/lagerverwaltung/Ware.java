package com.lagerverwaltung.taamefl2.lagerverwaltung;

public class Ware {
    private String name;
    private int number;
    private int threshold;
    private Boolean available;
    private Boolean toBuy;

    public Ware(String name, int number, int threshold ){
        this.name = name;
        this.number = number;
        this.threshold = threshold;
        this.available = checkIfAvailable(number);
        this.toBuy = checkfToBuy(threshold);
    }

    private Boolean checkfToBuy(int threshold) {
        return number == threshold;
    }

    private Boolean checkIfAvailable(int number) {
        return number != 0;
    }

}
