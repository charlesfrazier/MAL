use calamine::{open_workbook, Xlsx, Reader};
use std::collections::HashMap;
fn main() {
    let path = "/Users/cf/Downloads/276565941_按序号_宿舍分配问卷（前置）_4_3.xlsx";
    let mut workbook: Xlsx<_> = open_workbook(path).unwrap();
    let range = workbook.worksheet_range("Sheet1").unwrap();
    let mut map: HashMap<String, Vec<(i32, i32)>> = HashMap::new();
    for (i, row) in range.rows().skip(1).enumerate() {
        let key = row.iter().skip(6).take(10).map(|cell| cell.to_string()).collect::<String>();
        let score = row.iter().skip(16).filter_map(|cell| cell.to_string().parse::<i32>().ok()).sum();
        map.entry(key).or_insert_with(Vec::new).push((score, i as i32 + 1));
    }
    let mut final_allocation = Vec::new();
    let mut insufficient_keys = HashMap::new();
    for (key, mut scores) in map {
        if scores.len() < 4 {
            insufficient_keys.insert(key, scores);
        } else {
            scores.sort_by(|a, b| b.0.cmp(&a.0));
            while scores.len() >= 4 {
                let highest_score = scores.remove(0);
                let room: Vec<i32> = vec![highest_score.1].into_iter().chain(scores.iter().take(3).map(|&x| x.1)).collect();
                final_allocation.push((highest_score.0, room));
                scores = scores.split_off(3);
            }
            if !scores.is_empty() {
                insufficient_keys.insert(key, scores);
            }
        }
    }
    fn merge_keys(keys: HashMap<String, Vec<(i32, i32)>>, diff: usize) -> HashMap<String, Vec<(i32, i32)>> {
        let mut new_keys = HashMap::new();
        let mut visited = HashMap::new();
        for (key1, scores1) in &keys {
            if visited.contains_key(key1) {
                continue;
            }
            let mut merged_scores = scores1.clone();
            visited.insert(key1, true);
            for (key2, scores2) in &keys {
                if key1 != key2 && !visited.contains_key(key2) {
                    let differences = key1.chars().zip(key2.chars()).filter(|(a, b)| a != b).count();
                    if differences == diff {
                        merged_scores.extend(scores2.clone());
                        visited.insert(key2, true);
                    }
                }
            }
            new_keys.insert(key1.clone(), merged_scores);
        }
        new_keys
    }
    let mut diff = 1;
    while !insufficient_keys.is_empty() && diff <= 10 {
        insufficient_keys = merge_keys(insufficient_keys, diff);
        let mut new_insufficient_keys = HashMap::new();
        for (key, mut scores) in insufficient_keys {
            if scores.len() >= 4 {
                scores.sort_by(|a, b| b.0.cmp(&a.0));
                while scores.len() >= 4 {
                    let highest_score = scores.remove(0);
                    let room: Vec<i32> = vec![highest_score.1].into_iter().chain(scores.iter().take(3).map(|&x| x.1)).collect();
                    final_allocation.push((highest_score.0, room));
                    scores = scores.split_off(3);
                }
            }
            if !scores.is_empty() {
                new_insufficient_keys.insert(key, scores);
            }
        }
        insufficient_keys = new_insufficient_keys;
        diff += 1;
    }
    for allocation in final_allocation {
        println!("{:?}", allocation);
    }
    println!("Insufficient keys:");
    for (key, scores) in insufficient_keys {
        println!("{:?}: {:?}", key, scores);
    }
}
