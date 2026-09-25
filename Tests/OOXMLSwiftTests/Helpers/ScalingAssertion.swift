// ScalingAssertion.swift
// PsychQuant/ooxml-swift#174 — 以「規模比例」而非「絕對秒數」斷言線性時間。
//
// 過去的寫法是 `XCTAssertLessThan(Date().timeIntervalSince(started), 5.0)`：
// 機器負載一高（load average 900 時同一個測試量到 7.86 s），就和演算法複雜度
// 無關地失敗。這裡改成在 n 與 factor·n 兩個規模上交錯各量 `repeats` 次，取
// 中位數的比值：負載對兩個規模的影響大致相同，比值因此穩定；真正的平方時間
// 回歸仍會讓比值遠超過門檻。
//
// 門檻（factor = 4 時為 8）取線性比值 4 與平方比值 16 的幾何中點：兩側各留
// 兩倍餘裕。factor = 2（線性 2、平方 4、門檻約 3）兩側只有 1.3–1.5 倍餘裕，
// 所以預設用 4。固定成本（建 package、壓縮、讀檔）會把兩種比值都往 1 拉，
// 所以輸入在 `prepare` 裡建好、不計時，只量真正宣稱線性的那一段；呼叫端還要
// 挑「平方項在回歸時必然主導」的規模（每個呼叫點旁寫了依據）。

import Foundation
import XCTest

enum ScalingProbe {

    struct Measurement {
        let baseSize: Int
        let factor: Int
        let small: [TimeInterval]
        let large: [TimeInterval]

        var smallMedian: TimeInterval { Self.median(small) }
        var largeMedian: TimeInterval { Self.median(large) }
        var ratio: Double { largeMedian / max(smallMedian, .leastNonzeroMagnitude) }

        var summary: String {
            let fmt = { (xs: [TimeInterval]) in xs.map { String(format: "%.4f", $0) }.joined(separator: ", ") }
            return "n=\(baseSize): [\(fmt(small))] s; \(factor)n=\(baseSize * factor): [\(fmt(large))] s; "
                + String(format: "median ratio %.2f", ratio)
        }

        static func median(_ xs: [TimeInterval]) -> TimeInterval {
            let sorted = xs.sorted()
            guard !sorted.isEmpty else { return 0 }
            let mid = sorted.count / 2
            return sorted.count % 2 == 1 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2
        }
    }

    /// 單調時鐘（不受系統時間調整影響）量一次 `body` 的耗時。
    static func time(_ body: () throws -> Void) rethrows -> TimeInterval {
        let start = DispatchTime.now().uptimeNanoseconds
        try body()
        return Double(DispatchTime.now().uptimeNanoseconds - start) / 1e9
    }

    /// 兩個規模的輸入各 `prepare` 一次（不計時），先在 n 暖身一次（regex 編譯、
    /// lazy static 等一次性成本不算進量測），再以 n、factor·n 交錯量測，每輪
    /// 交換先後順序抵消負載漂移。
    static func measure<Input>(baseSize: Int, factor: Int, repeats: Int,
                               prepare: (Int) throws -> Input,
                               teardown: (Input) -> Void,
                               _ work: (Input) throws -> Void) rethrows -> Measurement {
        let smallInput = try prepare(baseSize)
        defer { teardown(smallInput) }
        let largeInput = try prepare(baseSize * factor)
        defer { teardown(largeInput) }
        try work(smallInput)
        var small: [TimeInterval] = [], large: [TimeInterval] = []
        for round in 0..<repeats {
            if round % 2 == 0 {
                small.append(try time { try work(smallInput) })
                large.append(try time { try work(largeInput) })
            } else {
                large.append(try time { try work(largeInput) })
                small.append(try time { try work(smallInput) })
            }
        }
        return Measurement(baseSize: baseSize, factor: factor, small: small, large: large)
    }
}

/// 斷言 `work` 對輸入規模近似線性：factor·n 與 n 的中位數耗時比低於 `maxRatio`。
///
/// `noiseFloor`：factor·n 的中位數低於這個秒數時不判比值——那個量級的量測被
/// 排程雜訊主導（一次 10 ms 的搶占就足以讓 1 ms 對 4 ms 的比值翻倍），也代表
/// 在這組規模下沒有可觀察的超線性成本。呼叫端選的規模必須讓平方時間的回歸
/// 遠高於這個下限。
func XCTAssertScalesLinearly<Input>(
    _ label: @autoclosure () -> String = "",
    baseSize: Int,
    factor: Int = 4,
    repeats: Int = 5,
    maxRatio: Double = 8,
    noiseFloor: TimeInterval = 0.05,
    file: StaticString = #filePath,
    line: UInt = #line,
    prepare: (Int) throws -> Input,
    teardown: (Input) -> Void = { _ in },
    _ work: (Input) throws -> Void
) rethrows {
    let measurement = try ScalingProbe.measure(
        baseSize: baseSize, factor: factor, repeats: repeats,
        prepare: prepare, teardown: teardown, work)
    let prefix = label().isEmpty ? "" : label() + ": "
    print("[ScalingProbe] \(prefix)\(measurement.summary)")
    guard measurement.largeMedian >= noiseFloor else { return }
    XCTAssertLessThan(
        measurement.ratio, maxRatio,
        "\(prefix)growing the input \(factor)× must grow the time about \(factor)× (linear), "
            + "not \(factor * factor)× (quadratic) — \(measurement.summary)",
        file: file, line: line)
}
