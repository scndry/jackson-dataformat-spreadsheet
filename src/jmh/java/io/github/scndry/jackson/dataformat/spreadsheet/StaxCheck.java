package io.github.scndry.jackson.dataformat.spreadsheet;

import javax.xml.stream.XMLInputFactory;

public class StaxCheck {
    public static void main(String[] args) {
        XMLInputFactory factory = XMLInputFactory.newInstance();
        System.out.println("XMLInputFactory impl: " + factory.getClass().getName());
    }
}
