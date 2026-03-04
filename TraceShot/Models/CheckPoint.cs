using System.Windows;

public class CheckPoint
{
    public string? Time { get; set; }    // "00:12" などの形式
    public double Seconds { get; set; } // 再生ジャンプ用の数値
    public string? Note { get; set; }    // メモ（「バグ発生」など）
    public override string ToString() => $"📌 [{Time}] {Note}";
    public string? ImagePath { get; set; } // 追加：保存した画像のフルパス
    public List<Rect> MarkRects { get; set; } = new List<Rect>();
}