"""Check the binary-image handlers with temporary files and a local HTTP server.

Only imageCacheFolder is redirected in the compiled test helper. The download,
validation, error and cleanup handlers are the actual SMFMac.applescript code.
"""
from base64 import b64decode
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import json
import re
import struct
import zlib
import subprocess
import tempfile
import threading
import unittest

HELPER = Path(__file__).resolve().parents[1] / "SMFMac.applescript"
PNG = b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=")


def rgb_png(width=128, height=64):
    def chunk(name, body):
        return struct.pack(">I", len(body)) + name + body + struct.pack(">I", zlib.crc32(name + body) & 0xFFFFFFFF)
    rows = []
    for y in range(height):
        rows.append(b"\0" + b"".join(bytes((x % 256, y % 256, (x + y) % 256)) for x in range(width)))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"".join(rows))) + chunk(b"IEND", b""))


class ImageServer(BaseHTTPRequestHandler):
    def log_message(self, *_):
        pass

    def do_GET(self):
        if self.path == "/redirect":
            self.send_response(302)
            self.send_header("Location", "/valid.png")
            self.end_headers()
            return
        if self.path == "/redirect-file":
            self.send_response(302)
            self.send_header("Location", "file:///etc/hosts")
            self.end_headers()
            return
        if self.path.startswith(("/valid", "/brackets")):
            status, body, content_type = 200, PNG, "image/png"
        elif self.path == "/not-image":
            status, body, content_type = 200, b"<html>Not an image</html>", "image/png"
        elif self.path == "/oversize":
            status, body, content_type = 200, PNG + b"x" * 10485760, "image/png"
        elif self.path == "/corrupt":
            status, body, content_type = 200, PNG[:24], "image/png"
        else:
            status, body, content_type = 404, b"Not found", "text/plain"
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass


class ImageTransportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory(prefix="smf-image-tests-")
        cls.folder = Path(cls.temp.name)
        cls.cache = cls.folder / "private-image-cache"
        source = HELPER.read_text()
        source, count = re.subn(r"on imageCacheFolder\(\).*?end imageCacheFolder",
                               "on imageCacheFolder()\n    return " + json.dumps(str(cls.cache)) + "\nend imageCacheFolder",
                               source, count=1, flags=re.S)
        assert count == 1
        test_source = cls.folder / "SMFMac.applescript"
        test_source.write_text(source)
        cls.compiled = cls.folder / "SMFMac.scpt"
        subprocess.run(["/usr/bin/osacompile", "-o", str(cls.compiled), str(test_source)],
                       check=True, capture_output=True, text=True, timeout=15)
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), ImageServer)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.base = "http://127.0.0.1:" + str(cls.server.server_port)

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.temp.cleanup()

    def handler(self, name, value):
        script = ("set helper to load script POSIX file " + json.dumps(str(self.compiled))
                  + "\nreturn helper's " + name + "(" + json.dumps(value, ensure_ascii=False) + ")")
        reply = subprocess.run(["/usr/bin/osascript", "-e", script],
                               capture_output=True, text=True, timeout=35, check=True)
        return reply.stdout.removesuffix("\n")

    def download(self, path):
        url = self.base + path if path.startswith("/") else path
        parts = self.handler("downloadImage", url).split("\n", 3)
        self.assertEqual(parts[0], "SMFIMAGE/1")
        self.assertEqual(len(parts), 4)
        return int(parts[1]), parts[2], parts[3]

    def prepare_local(self, path):
        parts = self.handler("prepareLocalImage", str(path)).split("\n", 3)
        self.assertEqual(parts[0], "SMFIMAGE/1")
        self.assertEqual(len(parts), 4)
        return int(parts[1]), parts[2], parts[3]

    def test_local_png_is_copied_decoded_and_original_preserved(self):
        source = self.folder / "local pixel fixture.png"
        source.write_bytes(PNG)
        status, path, error = self.prepare_local(source)
        self.assertEqual((status, error), (200, ""))
        self.assertEqual(Path(path).parent.parent, self.cache)
        self.assertTrue(Path(path).read_bytes().startswith(b"\x89PNG\r\n\x1a\n"))
        self.assertEqual(self.handler("releaseImage", path), "removed")
        self.assertEqual(source.read_bytes(), PNG)
        self.assert_cache_empty()

    def test_local_rgb_picture_dimensions_and_source_bytes_are_preserved(self):
        source = self.folder / "rgb chart fixture.png"
        original = rgb_png()
        source.write_bytes(original)
        status, path, error = self.prepare_local(source)
        self.assertEqual((status, error), (200, ""))
        dimensions = subprocess.run(["/usr/bin/sips", "-g", "pixelWidth", "-g", "pixelHeight", path],
                                    check=True, capture_output=True, text=True, timeout=10).stdout
        self.assertIn("pixelWidth:128", dimensions.replace(" ", ""))
        self.assertIn("pixelHeight:64", dimensions.replace(" ", ""))
        self.assertEqual(self.handler("releaseImage", path), "removed")
        self.assertEqual(source.read_bytes(), original)
        self.assert_cache_empty()

    def test_local_quoted_unicode_file_path_is_literal(self):
        source = self.folder / "chart's résumé $(printf injection).png"
        original = rgb_png(64, 32)
        source.write_bytes(original)
        status, path, error = self.prepare_local(source)
        self.assertEqual((status, error), (200, ""))
        self.assertTrue(Path(path).is_file())
        self.assertEqual(self.handler("releaseImage", path), "removed")
        self.assertEqual(source.read_bytes(), original)
        self.assert_cache_empty()

    def test_local_nonimage_and_corrupt_image_are_rejected(self):
        for name, data in (("not-image.png", b"not an image"), ("corrupt.png", PNG[:24])):
            source = self.folder / name
            source.write_bytes(data)
            status, path, error = self.prepare_local(source)
            self.assertEqual((status, path), (0, ""))
            self.assertTrue(error)
            self.assertEqual(source.read_bytes(), data)
        self.assert_cache_empty()

    def test_local_missing_relative_and_directory_paths_are_rejected(self):
        for source in (self.folder / "missing.png", "relative.png", self.folder):
            status, path, error = self.prepare_local(source)
            self.assertEqual((status, path), (0, ""))
            self.assertTrue(error)
        self.assert_cache_empty()

    def test_local_oversize_image_is_rejected_before_copy(self):
        source = self.folder / "oversize.png"
        source.write_bytes(PNG + b"x" * 10485760)
        status, path, error = self.prepare_local(source)
        self.assertEqual((status, path), (0, ""))
        self.assertIn("10 MiB", error)
        self.assert_cache_empty()

    def assert_cache_empty(self):
        self.assertEqual(list(self.cache.glob("image.*")), [])

    def test_png_is_decoded_and_release_removes_only_request_files(self):
        status, path, error = self.download("/valid.png")
        self.assertEqual((status, error), (200, ""))
        image = Path(path)
        self.assertTrue(image.is_file())
        self.assertTrue(image.read_bytes().startswith(b"\x89PNG\r\n\x1a\n"))
        self.assertEqual(image.parent.parent, self.cache)
        self.assertEqual(self.handler("releaseImage", path), "removed")
        self.assertFalse(image.parent.exists())
        self.assert_cache_empty()

    def test_image_redirect(self):
        status, path, error = self.download("/redirect")
        self.assertEqual((status, error), (200, ""))
        self.assertEqual(self.handler("releaseImage", path), "removed")
        self.assert_cache_empty()

    def test_brackets_are_literal_url_characters(self):
        status, path, error = self.download("/brackets?chart=AAPL,uu[h,a]dac")
        self.assertEqual((status, error), (200, ""))
        self.assertEqual(self.handler("releaseImage", path), "removed")

    def test_nonimage_http200_is_rejected_and_cleaned(self):
        status, path, error = self.download("/not-image")
        self.assertEqual((status, path), (200, ""))
        self.assertIn("not a PNG", error)
        self.assert_cache_empty()

    def test_404_is_rejected_and_cleaned(self):
        status, path, error = self.download("/missing")
        self.assertEqual((status, path), (404, ""))
        self.assertIn("HTTP 404", error)
        self.assert_cache_empty()

    def test_corrupt_png_is_rejected_and_cleaned(self):
        _, path, error = self.download("/corrupt")
        self.assertEqual(path, "")
        self.assertTrue(error)
        self.assert_cache_empty()

    def test_oversize_image_is_rejected_and_cleaned(self):
        _, path, error = self.download("/oversize")
        self.assertEqual(path, "")
        self.assertTrue(error)
        self.assert_cache_empty()

    def test_nonhttp_protocol_and_newline_are_rejected(self):
        for url in ("file:///etc/hosts", self.base + "/valid.png\nextra"):
            status, path, error = self.download(url)
            self.assertEqual((status, path), (0, ""))
            self.assertTrue(error)
        self.assert_cache_empty()

    def test_redirect_cannot_read_a_local_file(self):
        _, path, error = self.download("/redirect-file")
        self.assertEqual(path, "")
        self.assertTrue(error)
        self.assert_cache_empty()

    def test_release_refuses_external_or_traversed_paths(self):
        outside = self.folder / "outside.png"
        outside.write_bytes(PNG)
        for path in (str(outside), str(self.cache / "image.abc/../../outside.png")):
            self.assertEqual(self.handler("releaseImage", path), "ignored")
        self.assertEqual(outside.read_bytes(), PNG)

    def test_release_requires_request_ownership_marker(self):
        folder = self.cache / "image.notours"
        folder.mkdir(parents=True)
        image = folder / "image.png"
        image.write_bytes(PNG)
        self.assertEqual(self.handler("releaseImage", str(image)), "unavailable")
        self.assertEqual(image.read_bytes(), PNG)
        image.unlink()
        folder.rmdir()

    def test_release_preserves_unrecognized_files(self):
        _, path, error = self.download("/valid.png")
        self.assertEqual(error, "")
        folder = Path(path).parent
        extra = folder / "unrelated.txt"
        extra.write_text("keep")
        self.assertEqual(self.handler("releaseImage", path), "unavailable")
        self.assertEqual(extra.read_text(), "keep")
        extra.unlink()
        folder.rmdir()


if __name__ == "__main__":
    unittest.main(verbosity=2)
