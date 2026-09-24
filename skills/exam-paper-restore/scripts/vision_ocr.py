#!/usr/bin/env python3
"""macOS Vision OCR for exam paper page images.

Usage:
    python vision_ocr.py page1.jpg page2.jpg ...          # full-page OCR, sorted top-to-bottom
    python vision_ocr.py page.jpg 0.44 0.46               # OCR only rows in [0.44, 0.46] of page height
    python vision_ocr.py page.jpg 0.44 0.46 0.0 0.6       # also restrict x to [0.0, 0.6]

Requires: pyobjc-framework-Vision, pyobjc-framework-Quartz, Pillow
    pip install pyobjc-framework-Vision pyobjc-framework-Quartz Pillow

Output: one line per detected text block, prefixed with its vertical position (0=top, 1=bottom),
        so you can map text back to page coordinates.
"""
import os
import sys
import tempfile

from PIL import Image
import Quartz
from Foundation import NSURL
import Vision


def ocr_file(path, scale=2.0, languages=("zh-Hans", "en-US"), correct=False):
    url = NSURL.fileURLWithPath_(path)
    src = Quartz.CGImageSourceCreateWithURL(url, None)
    img = Quartz.CGImageSourceCreateImageAtIndex(src, 0, None)
    w = Quartz.CGImageGetWidth(img)
    h = Quartz.CGImageGetHeight(img)
    cs = Quartz.CGColorSpaceCreateDeviceRGB()
    ctx = Quartz.CGBitmapContextCreate(
        None, int(w * scale), int(h * scale), 8, 0, cs,
        Quartz.kCGImageAlphaPremultipliedLast | Quartz.kCGBitmapByteOrder32Big)
    Quartz.CGContextSetInterpolationQuality(ctx, Quartz.kCGInterpolationHigh)
    Quartz.CGContextDrawImage(ctx, Quartz.CGRectMake(0, 0, w * scale, h * scale), img)
    big = Quartz.CGBitmapContextCreateImage(ctx)

    req = Vision.VNRecognizeTextRequest.alloc().init()
    req.setRecognitionLevel_(0)          # 0 = accurate
    req.setRecognitionLanguages_(list(languages))
    req.setUsesLanguageCorrection_(correct)
    handler = Vision.VNImageRequestHandler.alloc().initWithCGImage_options_(big, None)
    handler.performRequests_error_([req], None)

    out = []
    for r in req.results():
        cand = r.topCandidates_(1)[0]
        bb = r.boundingBox()
        out.append((1 - bb.origin.y, bb.origin.x, cand.string()))
    out.sort()
    return out


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return 1

    if len(args) >= 3 and args[1].replace(".", "", 1).isdigit():
        path = args[0]
        y0, y1 = float(args[1]), float(args[2])
        x0 = float(args[3]) if len(args) > 3 else 0.0
        x1 = float(args[4]) if len(args) > 4 else 1.0
        im = Image.open(path)
        w, h = im.size
        crop = im.crop((int(x0 * w), int(y0 * h), int(x1 * w), int(y1 * h)))
        tmp = tempfile.mktemp(suffix=".png")
        crop.save(tmp)
        try:
            for y, x, s in ocr_file(tmp, scale=3.0):
                print("%.3f %s" % (y, s))
        finally:
            os.unlink(tmp)
        return 0

    for path in args:
        print("########## " + path)
        for y, x, s in ocr_file(path):
            print("%.3f %s" % (y, s))
    return 0


if __name__ == "__main__":
    sys.exit(main())
