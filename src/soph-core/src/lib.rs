/// シンプルな挨拶文字列を返す
pub fn greeting() -> &'static str {
    "soph-core から こんにちは！"
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_greeting() {
        assert!(!greeting().is_empty());
    }
}
