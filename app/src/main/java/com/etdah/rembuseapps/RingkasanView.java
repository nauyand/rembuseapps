package com.etdah.rembuseapps;

import android.content.Context;
import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.Paint;
import android.graphics.RectF;
import android.util.AttributeSet;
import android.util.TypedValue;
import android.view.View;

import java.text.NumberFormat;
import java.util.Locale;

/**
 * Ringkasan visual: batang per jenis pengeluaran (porsi dari total)
 * dan satu batang pemakaian kas.
 */
public class RingkasanView extends View {

    private static final String[] LABEL = {"Material", "Jasa", "Operasional"};

    private static final int[] WARNA = {
            Color.parseColor("#6741A5"),
            Color.parseColor("#9270D5"),
            Color.parseColor("#B69AF5")
    };

    private static final int WARNA_JALUR = Color.parseColor("#EEE8F7");
    private static final int WARNA_TEKS = Color.parseColor("#40394B");
    private static final int WARNA_TEKS_REDUP = Color.parseColor("#8A8296");
    private static final int WARNA_KAS = Color.parseColor("#2E7D5B");
    private static final int WARNA_LEBIH = Color.parseColor("#B3261E");

    private final double[] nilai = new double[3];
    private double kas = 0;

    private final Paint paintTeks = new Paint(Paint.ANTI_ALIAS_FLAG);
    private final Paint paintTeksKanan = new Paint(Paint.ANTI_ALIAS_FLAG);
    private final Paint paintBatang = new Paint(Paint.ANTI_ALIAS_FLAG);
    private final RectF kotak = new RectF();

    private final NumberFormat rupiah =
            NumberFormat.getCurrencyInstance(new Locale("id", "ID"));

    private final float tinggiTeks;
    private final float tinggiBatang;
    private final float jarakBaris;
    private final float jarakTeksBatang;

    public RingkasanView(Context context) {
        this(context, null);
    }

    public RingkasanView(Context context, AttributeSet attrs) {
        super(context, attrs);

        rupiah.setMaximumFractionDigits(0);

        float sp13 = TypedValue.applyDimension(TypedValue.COMPLEX_UNIT_SP, 13,
                getResources().getDisplayMetrics());

        paintTeks.setTextSize(sp13);
        paintTeks.setColor(WARNA_TEKS);

        paintTeksKanan.setTextSize(sp13);
        paintTeksKanan.setColor(WARNA_TEKS_REDUP);
        paintTeksKanan.setTextAlign(Paint.Align.RIGHT);

        tinggiTeks = sp13 * 1.25f;
        tinggiBatang = dp(8);
        jarakTeksBatang = dp(5);
        jarakBaris = dp(12);
    }

    /** Memperbarui angka lalu menggambar ulang. */
    public void setData(double material, double jasa, double operasional, double kas) {
        nilai[0] = material;
        nilai[1] = jasa;
        nilai[2] = operasional;
        this.kas = kas;
        invalidate();
    }

    private float dp(float v) {
        return TypedValue.applyDimension(TypedValue.COMPLEX_UNIT_DIP, v,
                getResources().getDisplayMetrics());
    }

    private float tinggiSatuBaris() {
        return tinggiTeks + jarakTeksBatang + tinggiBatang;
    }

    @Override
    protected void onMeasure(int widthMeasureSpec, int heightMeasureSpec) {
        int lebar = MeasureSpec.getSize(widthMeasureSpec);
        // 3 baris jenis + 1 baris kas, dengan jarak ekstra sebelum baris kas.
        float tinggi = getPaddingTop() + getPaddingBottom()
                + 4 * tinggiSatuBaris() + 3 * jarakBaris + dp(6);
        setMeasuredDimension(lebar, (int) Math.ceil(tinggi));
    }

    @Override
    protected void onDraw(Canvas canvas) {
        super.onDraw(canvas);

        float kiri = getPaddingLeft();
        float kanan = getWidth() - getPaddingRight();
        float y = getPaddingTop();

        double total = nilai[0] + nilai[1] + nilai[2];

        for (int i = 0; i < 3; i++) {
            double porsi = total > 0 ? nilai[i] / total : 0;

            gambarBaris(canvas, kiri, kanan, y,
                    LABEL[i],
                    rupiah.format(nilai[i]) + "  •  " + Math.round(porsi * 100) + "%",
                    (float) porsi,
                    WARNA[i]);

            y += tinggiSatuBaris() + jarakBaris;
        }

        y += dp(6);

        // Pemakaian kas.
        double pakai = kas > 0 ? total / kas : (total > 0 ? 1 : 0);
        boolean lebih = kas > 0 ? total > kas : total > 0;

        String keterangan = kas > 0
                ? Math.round(pakai * 100) + "% dari " + rupiah.format(kas)
                : "kas belum diisi";

        gambarBaris(canvas, kiri, kanan, y,
                lebih ? "Kas terlampaui" : "Kas terpakai",
                keterangan,
                (float) Math.min(pakai, 1),
                lebih ? WARNA_LEBIH : WARNA_KAS);
    }

    private void gambarBaris(Canvas canvas, float kiri, float kanan, float y,
                             String label, String nilaiTeks, float porsi, int warna) {

        float garisDasar = y + tinggiTeks * 0.8f;
        canvas.drawText(label, kiri, garisDasar, paintTeks);
        canvas.drawText(nilaiTeks, kanan, garisDasar, paintTeksKanan);

        float atasBatang = y + tinggiTeks + jarakTeksBatang;
        float radius = tinggiBatang / 2f;

        paintBatang.setColor(WARNA_JALUR);
        kotak.set(kiri, atasBatang, kanan, atasBatang + tinggiBatang);
        canvas.drawRoundRect(kotak, radius, radius, paintBatang);

        if (porsi > 0) {
            float lebarIsi = Math.max((kanan - kiri) * porsi, tinggiBatang);
            paintBatang.setColor(warna);
            kotak.set(kiri, atasBatang, kiri + lebarIsi, atasBatang + tinggiBatang);
            canvas.drawRoundRect(kotak, radius, radius, paintBatang);
        }
    }
}
