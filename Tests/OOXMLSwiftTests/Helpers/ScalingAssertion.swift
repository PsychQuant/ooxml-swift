// ScalingAssertion.swift
// PsychQuant/ooxml-swift#174 — 以「規模比例」而非「絕對秒數」斷言線性時間。
//
// 過去的寫法是 `XCTAssertLessThan(Date().timeIntervalSince(started), 5.0)`：
// 機器負載一高（load average 900 時同一個測試量到 7.86 s），就和演算法複雜度
// 無關地失敗。這裡改成在 n 與 factor·n 兩個規模上交錯各量 `repeats` 次，
// 比較兩個規模的最小耗時：負載不改變演算法的成長率，真正的平方時間回歸仍會
// 讓比值遠超過門檻。
//
// 門檻（factor = 4 時為 8）取線性比值 4 與平方比值 16 的幾何中點：兩側各留
// 兩倍餘裕。factor = 2（線性 2、平方 4、門檻約 3）兩側只有 1.3–1.5 倍餘裕，
// 所以預設用 4。固定成本（建 package、壓縮、讀檔）會把兩種比值都往 1 拉，
// 所以輸入在 `prepare` 裡建好、不計時，只量真正宣稱線性的那一段；呼叫端還要
// 挑「平方項在回歸時必然主導」的規模（每個呼叫點旁寫了依據）。
//
// 獨立審查後的修正（revooxmlc HIGH-1）：第一版量牆鐘時間、取中位數，在 16 個
// 並行 xctest 行程、load 45–58 下 720 次比例斷言誤報 5 次——n 端樣本只有
// 15–20 ms，排程搶占與執行時間成正比，4n 端吸收的干擾較多，比值被系統性地往
// 上推。現在：
// - 量**本執行緒的 CPU 時間**（`CLOCK_THREAD_CPUTIME_ID`）：被搶占、等待
//   排程的時間不算進去。被量的工作必須在呼叫執行緒上同步完成——目前的呼叫點
//   （`PackageInspector`、`DocxWriter`、`DocxReader`）都沒有派工到其他執行緒
//   （grep 過 `DispatchQueue`／`concurrentPerform`／`Task`／`Process`）；若將來
//   改成並行，thread CPU 時間會少算。這個前提**由程式強制**，不只寫在註解裡：
//   每次量測也記錄 process CPU，工作若跑到其他執行緒（process CPU 超過 thread
//   CPU 的 `offThreadLimit` 倍），斷言直接失敗並說明原因，而不是用少算的 thread
//   CPU 判定、默默通過。第二輪審查（revooxmld 的注入 C）證實了這個盲點：把一段
//   平方時間的工作搬到 `DispatchQueue.global()` 上執行，thread CPU 比值 4.46、
//   測試照樣通過，牆鐘比值卻是 10–16、process CPU 441 s 對 thread CPU 5.8 s。
//   正常執行時兩者比值是 1.000（16 個並行 xctest 行程、load 186–212 實測）。
//   選 thread 而非 process：兩者在 16 個並行行程、load 73–110 下都是 0 誤報
//   （各 384 次量測），但 process CPU 時間會把同一行程內其他執行緒也算進去——
//   實測同一段工作在同行程有 3 條忙碌執行緒時，process CPU 10.46 s、thread CPU
//   2.51 s（安靜時 2.23 s／2.26 s）。Swift Testing 會在同一行程內並行跑測試。
// - 統計量用**最小值**：干擾只會讓一次量測變長，不會變短，最小值最接近工作
//   本身的成本（P/E 核心、快取冷熱都一樣）。
// - 會被判定的呼叫點把 n 加大到小端約 100 ms，讓計時解析度與殘餘雜訊相對可忽略。

import Darwin
import Foundation
import XCTest

enum ScalingProbe {

    struct Measurement {
        let baseSize: Int
        let factor: Int
        /// 本執行緒 CPU 時間（秒）——判定用。
        let small: [TimeInterval]
        let large: [TimeInterval]
        /// 同一次量測的牆鐘時間（秒）——只印在摘要裡供診斷，不參與判定。
        let smallWall: [TimeInterval]
        let largeWall: [TimeInterval]
        /// 同一次量測的行程 CPU 時間（秒）——只用來偵測工作是否跑到其他執行緒。
        let smallProcess: [TimeInterval]
        let largeProcess: [TimeInterval]

        var smallMin: TimeInterval { small.min() ?? 0 }
        var largeMin: TimeInterval { large.min() ?? 0 }
        var ratio: Double { largeMin / max(smallMin, .leastNonzeroMagnitude) }
        /// 行程 CPU 總和 ÷ 本執行緒 CPU 總和。工作全在呼叫執行緒上時是 1。
        var offThreadFactor: Double {
            let process = (smallProcess + largeProcess).reduce(0, +)
            let thread = (small + large).reduce(0, +)
            return process / max(thread, .leastNonzeroMagnitude)
        }
        var largeProcessMin: TimeInterval { largeProcess.min() ?? 0 }

        var summary: String {
            let fmt = { (xs: [TimeInterval]) in xs.map { String(format: "%.4f", $0) }.joined(separator: ", ") }
            return "thread CPU n=\(baseSize): [\(fmt(small))] s; \(factor)n=\(baseSize * factor): [\(fmt(large))] s; "
                + String(format: "min ratio %.2f", ratio)
                + " (wall: [\(fmt(smallWall))] / [\(fmt(largeWall))] s; "
                + String(format: "process/thread CPU %.2f)", offThreadFactor)
        }
    }

    /// 量一次 `body`：回傳（本執行緒 CPU 時間, 牆鐘時間, 行程 CPU 時間），單位秒。
    static func time(_ body: () throws -> Void) rethrows
        -> (cpu: TimeInterval, wall: TimeInterval, process: TimeInterval) {
        let process0 = clock_gettime_nsec_np(CLOCK_PROCESS_CPUTIME_ID)
        let cpu0 = clock_gettime_nsec_np(CLOCK_THREAD_CPUTIME_ID)
        let wall0 = clock_gettime_nsec_np(CLOCK_UPTIME_RAW)
        try body()
        let cpu1 = clock_gettime_nsec_np(CLOCK_THREAD_CPUTIME_ID)
        let wall1 = clock_gettime_nsec_np(CLOCK_UPTIME_RAW)
        let process1 = clock_gettime_nsec_np(CLOCK_PROCESS_CPUTIME_ID)
        return (Double(cpu1 - cpu0) / 1e9, Double(wall1 - wall0) / 1e9,
                Double(process1 - process0) / 1e9)
    }

    /// 兩個規模的輸入各 `prepare` 一次（不計時），先在 n 暖身一次（regex 編譯、
    /// lazy static 等一次性成本不算進量測），再以 n、factor·n 交錯量測，每輪
    /// 交換先後順序。
    static func measure<Input>(baseSize: Int, factor: Int, repeats: Int,
                               prepare: (Int) throws -> Input,
                               teardown: (Input) -> Void,
                               _ work: (Input) throws -> Void) rethrows -> Measurement {
        let smallInput = try prepare(baseSize)
        defer { teardown(smallInput) }
        let largeInput = try prepare(baseSize * factor)
        defer { teardown(largeInput) }
        try work(smallInput)
        var small: [(cpu: TimeInterval, wall: TimeInterval, process: TimeInterval)] = []
        var large: [(cpu: TimeInterval, wall: TimeInterval, process: TimeInterval)] = []
        for round in 0..<repeats {
            if round % 2 == 0 {
                small.append(try time { try work(smallInput) })
                large.append(try time { try work(largeInput) })
            } else {
                large.append(try time { try work(largeInput) })
                small.append(try time { try work(smallInput) })
            }
        }
        return Measurement(baseSize: baseSize, factor: factor,
                           small: small.map(\.cpu), large: large.map(\.cpu),
                           smallWall: small.map(\.wall), largeWall: large.map(\.wall),
                           smallProcess: small.map(\.process), largeProcess: large.map(\.process))
    }
}

/// 斷言 `work` 對輸入規模近似線性：factor·n 與 n 的最小 CPU 耗時比低於 `maxRatio`。
///
/// `noiseFloor`：factor·n 的最小 CPU 耗時低於這個秒數時不判比值——代表在這組
/// 規模下根本沒有可量到的規模相關成本（例如 `PackageInspector` 的病態 payload：
/// 線性版本主要是固定的解壓成本，4n 也只有十幾毫秒）。這類呼叫點刻意不加大 n：
/// 它們防的是 35–82 s 級的回歸，回歸時兩端都是秒級、一定會被判定；把 n 加大到
/// 線性時也超過下限，只會讓回歸變成數小時的 hang。會在正常執行中被判定的呼叫點，
/// n 端應在 100 ms 量級。
func XCTAssertScalesLinearly<Input>(
    _ label: @autoclosure () -> String = "",
    baseSize: Int,
    factor: Int = 4,
    repeats: Int = 5,
    maxRatio: Double = 8,
    noiseFloor: TimeInterval = 0.05,
    offThreadLimit: Double = 1.5,
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
    // 先檢查工作是否留在呼叫執行緒上，再看雜訊下限：工作搬到其他執行緒時，
    // thread CPU 會變小，甚至低於下限而被當成「不判定」——那正是要擋的盲點。
    if measurement.largeProcessMin >= noiseFloor && measurement.offThreadFactor > offThreadLimit {
        XCTFail(
            "\(prefix)the measured work ran off the calling thread "
                + String(format: "(process/thread CPU %.2f > %.2f)", measurement.offThreadFactor, offThreadLimit)
                + ", so thread CPU time no longer measures it and cannot tell linear from quadratic. "
                + "Keep the measured work synchronous on the calling thread, or measure it with "
                + "process CPU time in a process where nothing else runs — \(measurement.summary)",
            file: file, line: line)
        return
    }
    let judged = measurement.largeMin >= noiseFloor
    print("[ScalingProbe] \(prefix)\(measurement.summary)\(judged ? "" : " — below noise floor, not judged")")
    guard judged else { return }
    XCTAssertLessThan(
        measurement.ratio, maxRatio,
        "\(prefix)growing the input \(factor)× must grow the CPU time about \(factor)× (linear), "
            + "not \(factor * factor)× (quadratic) — \(measurement.summary)",
        file: file, line: line)
}
