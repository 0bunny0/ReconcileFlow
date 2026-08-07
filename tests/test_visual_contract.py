from hashlib import sha256
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class VisualContractTests(unittest.TestCase):
    def test_supplied_index_html_is_byte_identical(self) -> None:
        content = (ROOT / "assets" / "index.html").read_bytes()
        self.assertEqual(
            sha256(content).hexdigest(),
            "66fefa48ba7bb515cd88534f4eeaa85051104323c1855a3a8bd1a6481a2a3bbc",
        )

    def test_approved_styles_css_is_byte_identical(self) -> None:
        content = (ROOT / "assets" / "styles.css").read_bytes()
        self.assertEqual(
            sha256(content).hexdigest(),
            "0a0936f5e9eed88c31c3baaa44543ba7fa51fe57760d1a04b4d917ab0a4c3f50",
        )


if __name__ == "__main__":
    unittest.main()
