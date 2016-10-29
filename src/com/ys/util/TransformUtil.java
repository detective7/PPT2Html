package com.ys.util;

public class TransformUtil {
    // RGB转16进制色
    public static String toHex(String rgb) {
        String all = rgb.substring(rgb.indexOf('[') + 1, rgb.indexOf(']'));
        String[] alls, rs, gs, bs;
        alls = all.split(",");
        rs = alls[0].split("=");
        gs = alls[1].split("=");
        bs = alls[2].split("=");
        int r, g, b;
        r = Integer.valueOf(rs[1]);
        g = Integer.valueOf(gs[1]);
        b = Integer.valueOf(bs[1]);

        return "#" + toBrowserHexValue(r) + toBrowserHexValue(g) + toBrowserHexValue(b);
    }

    private static String toBrowserHexValue(int number) {
        StringBuilder builder = new StringBuilder(Integer.toHexString(number & 0xff));
        while (builder.length() < 2) {
            builder.append("0");
        }
        return builder.toString().toUpperCase();
    }
}
