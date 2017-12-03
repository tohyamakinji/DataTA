package main;

import org.json.simple.parser.ParseException;
import writer.ExcelWriter;

import java.io.IOException;

public class DataProgram {

    private void execute(byte CODE) throws IOException, ParseException {
        switch (CODE) {
            case 1 :
                writeTest();
                break;
            case 2 :
                writeTraining();
                break;
            case 3 :
                writeTesting();
                break;
            case 4 :
                writeSenseTAG();
                break;
            case 5 :
                writeFeatures();
                break;
            default :
                writeTest();
                break;
        }
    }

    private void writeFeatures() throws IOException, ParseException {
        ExcelWriter writer = new ExcelWriter();
        writer.features();
    }

    private void writeSenseTAG() throws IOException, ParseException {
        ExcelWriter writer = new ExcelWriter();
        writer.senseTAG();
    }

    private void writeTesting() throws IOException, ParseException {
        ExcelWriter writer = new ExcelWriter();
        writer.testing();
    }

    private void writeTraining() throws IOException, ParseException {
        ExcelWriter writer = new ExcelWriter();
        writer.training();
    }

    private void writeTest() throws IOException {
        ExcelWriter writer = new ExcelWriter();
        writer.test();
    }

    public static void main(String[] args) {
        System.out.println("PLEASE WAIT !!");
        try {
            byte CODE = 5;
            DataProgram program = new DataProgram();
            program.execute(CODE);
            System.out.println("DONE !!");
        } catch (Exception e) {
            System.out.println("ERROR !!");
            e.printStackTrace();
        }
    }

}