package com.example.demo.func_test;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

public class CircleTest {

    @Test
    public void givenRadius_whenCalculateArea_thenReturnArea() {
        double actualArea = Circle.calculateArea(2d);
        double expectedArea = 2 * 2 * Math.PI;
        Assertions.assertEquals(expectedArea, actualArea);
    }

}
