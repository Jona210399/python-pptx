import argparse
from pathlib import Path

from PIL import Image


def generate_images(
    count: int = 10, size: tuple = (256, 256), out_dir: Path = Path("images")
) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    colors: list = [
        (255, 0, 0),
        (0, 255, 0),
        (0, 0, 255),
        (255, 255, 0),
        (255, 0, 255),
        (0, 255, 255),
        (128, 0, 128),
        (255, 165, 0),
        (0, 128, 0),
        (128, 128, 128),
    ]
    for i in range(count):
        color = colors[i % len(colors)]
        img = Image.new("RGB", size, color)
        path = out_dir / f"color_{i + 1:02d}.png"
        img.save(path, "PNG")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--count", type=int, default=10)
    parser.add_argument("--width", type=int, default=256)
    parser.add_argument("--height", type=int, default=256)
    parser.add_argument("--out", type=str, default="images")
    args = parser.parse_args()
    generate_images(count=args.count, size=(args.width, args.height), out_dir=Path(args.out))


if __name__ == "__main__":
    main()
