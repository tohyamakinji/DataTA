package processing;

import org.json.simple.JSONArray;

import java.util.ArrayList;
import java.util.List;

public class Collocation {

    private JSONArray arrayConjunction;

    public Collocation(JSONArray arrayConjunction) {
        this.arrayConjunction = arrayConjunction;
    }

    public List<String> process(List<String> splitText, String word) {
        List<Integer> posLocation = new ArrayList<>();
        List<String> wordBag = new ArrayList<>();

        for (int i=0; i < splitText.size(); i++) {
            if (splitText.get(i).equals(word))
                posLocation.add(i);
        }
        int count = 0, x, y, findX, findY;
        boolean isFinishLeft, isFinishRight;
        while (count < posLocation.size()) {
            x = posLocation.get(count) - 1;
            y = posLocation.get(count) + 1;
            findX = 0;
            findY = 0;
            isFinishLeft = false;
            isFinishRight = false;
            while (true) {
                if (x > -1) {
                    if (!splitText.get(x).equals(word) && !wordBag.contains(splitText.get(x))) {
                        if (!isConjunction(splitText.get(x))) {
                            wordBag.add(splitText.get(x));
                            findX++;
                        }
                    }
                    x--;
                    if (findX == 3)
                        x = -1; // end loop
                } else
                    isFinishLeft = true;
                if (y < splitText.size()) {
                    if (!splitText.get(y).equals(word) && !wordBag.contains(splitText.get(y))) {
                        if (!isConjunction(splitText.get(y))) {
                            wordBag.add(splitText.get(y));
                            findY++;
                        }
                    }
                    y++;
                    if (findY == 3)
                        y = splitText.size(); // end loop
                } else
                    isFinishRight = true;
                if (isFinishLeft && isFinishRight)
                    break;
            }
            count++;
        }
        return wordBag;
    }

    private boolean isConjunction(String s) {
        boolean status = false;
        for (Object o : arrayConjunction) {
            if (o.equals(s)) {
                status = true;
            }
        }
        return status;
    }

}