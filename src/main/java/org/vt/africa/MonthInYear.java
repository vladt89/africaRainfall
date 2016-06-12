package org.vt.africa;

/**
 * @author vladimir.tikhomirov
 */
public class MonthInYear {
    private double sum;
    private double month;
    private double year;

    public MonthInYear(double month, double year, double sumForMonth) {
        this.month = month;
        this.year = year;
        this.sum = sumForMonth;
    }

    public double getSum() {
        return sum;
    }
}
