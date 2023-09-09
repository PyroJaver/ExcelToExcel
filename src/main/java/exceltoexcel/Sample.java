package exceltoexcel;

import java.util.HashMap;

public class Sample {
    public Sample(HashMap<String, String> sampleParts) {
        this.sampleParts = sampleParts;
    }

    public HashMap<String, String> getSampleParts() {
        return sampleParts;
    }

    public void setSampleParts(HashMap<String, String> sampleParts) {
        this.sampleParts = sampleParts;
    }

    private HashMap<String, String> sampleParts;
}
