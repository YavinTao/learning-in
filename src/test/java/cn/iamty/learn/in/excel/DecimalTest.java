package cn.iamty.learn.in.excel;

import org.junit.Test;

import java.text.DecimalFormat;

public class DecimalTest {


    @Test
    public void testString2Integer() {
        double a = 1;
        String af = String.format("%.2f", a);
        Double b = Double.valueOf(af);
        System.out.println(b);
    }
    @Test
    public void testDecimalFormat() {
        double a = 0.0867587567;
        DecimalFormat decimalFormat = new DecimalFormat("#.00");

        String b = decimalFormat.format(a);

        System.out.println(b);
    }

}
