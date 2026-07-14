# -*- coding: utf-8 -*-
"""
压缩 SVGA 文件。

SVGA 2.0 文件本质是 zlib 压缩后的 protobuf(MovieEntity)，
体积几乎全部来自内嵌的 PNG 帧图(images 字段, field 3, map<string,bytes>)。
本脚本只重新压缩这些 PNG(可选缩放 + 调色板量化)，其余 protobuf 字段原样保留，
因此不改变动画结构/时长/坐标，只减小体积。

用法:
    python compress_svga.py 输入.svga [输出.svga] [--scale 0.7] [--colors 256]

默认: 输出为 <输入>_compressed.svga, scale=1.0(不缩放), colors=256(量化到256色)
"""
import sys
import io
import zlib
import argparse

from PIL import Image


# ---------- 最小 protobuf 读写 ----------
def read_varint(buf, i):
    shift = 0
    result = 0
    while True:
        b = buf[i]
        i += 1
        result |= (b & 0x7F) << shift
        if not (b & 0x80):
            break
        shift += 7
    return result, i


def write_varint(value):
    out = bytearray()
    while True:
        b = value & 0x7F
        value >>= 7
        if value:
            out.append(b | 0x80)
        else:
            out.append(b)
            break
    return bytes(out)


def ld(field_num, content):
    """length-delimited 字段: tag + varint(len) + content"""
    return write_varint((field_num << 3) | 2) + write_varint(len(content)) + content


# ---------- PNG 重压缩 ----------
def recompress_png(data, scale, colors):
    try:
        im = Image.open(io.BytesIO(data))
        im.load()
    except Exception:
        return data  # 非图片/无法解析，原样返回

    im = im.convert("RGBA")

    if scale != 1.0:
        w, h = im.size
        nw, nh = max(1, int(round(w * scale))), max(1, int(round(h * scale)))
        im = im.resize((nw, nh), Image.LANCZOS)

    out = io.BytesIO()
    if colors and colors <= 256:
        # 保留 alpha 的调色板量化
        alpha = im.split()[3]
        pal = im.convert("RGB").quantize(colors=min(colors, 256), method=Image.FASTOCTREE)
        pal = pal.convert("RGBA")
        pal.putalpha(alpha)
        pal = pal.quantize(colors=min(colors, 256), method=Image.FASTOCTREE)
        pal.save(out, format="PNG", optimize=True)
    else:
        im.save(out, format="PNG", optimize=True)

    new = out.getvalue()
    return new if len(new) < len(data) else data  # 变大则不替换


# ---------- 遍历顶层 MovieEntity，只改 field 3 ----------
def process(raw, scale, colors):
    i = 0
    n = len(raw)
    out = bytearray()
    count = 0
    while i < n:
        tag, i = read_varint(raw, i)
        field_num = tag >> 3
        wire_type = tag & 7

        if wire_type == 0:  # varint
            start = i
            _, i = read_varint(raw, i)
            out += write_varint(tag) + raw[start:i]
        elif wire_type == 1:  # 64-bit
            out += write_varint(tag) + raw[i:i + 8]
            i += 8
        elif wire_type == 5:  # 32-bit
            out += write_varint(tag) + raw[i:i + 4]
            i += 4
        elif wire_type == 2:  # length-delimited
            length, i = read_varint(raw, i)
            payload = raw[i:i + length]
            i += length
            if field_num == 3:  # images map entry
                entry = rebuild_image_entry(payload, scale, colors)
                if entry is not None:
                    count += 1
                    out += ld(3, entry)
                else:
                    out += ld(3, payload)
            else:
                out += ld(field_num, payload)
        else:
            raise ValueError("unsupported wire type %d at %d" % (wire_type, i))

    print("  重压缩帧图 %d 张" % count)
    return bytes(out)


def rebuild_image_entry(payload, scale, colors):
    """map entry: field1=key(string), field2=value(bytes=png)"""
    i = 0
    n = len(payload)
    key = None
    value = None
    while i < n:
        tag, i = read_varint(payload, i)
        fn = tag >> 3
        wt = tag & 7
        if wt != 2:
            return None  # 非预期结构，交回原样
        length, i = read_varint(payload, i)
        chunk = payload[i:i + length]
        i += length
        if fn == 1:
            key = chunk
        elif fn == 2:
            value = chunk
    if value is None:
        return None
    new_value = recompress_png(value, scale, colors)
    entry = bytearray()
    if key is not None:
        entry += ld(1, key)
    entry += ld(2, new_value)
    return bytes(entry)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--scale", type=float, default=1.0, help="帧图缩放比例, 如0.7")
    ap.add_argument("--colors", type=int, default=256, help="调色板颜色数, 0=不量化")
    args = ap.parse_args()

    out_path = args.output
    if not out_path:
        if args.input.lower().endswith(".svga"):
            out_path = args.input[:-5] + "_compressed.svga"
        else:
            out_path = args.input + "_compressed.svga"

    src = open(args.input, "rb").read()
    raw = zlib.decompress(src)
    print("原始文件: %.2f MB (protobuf %.2f MB)" % (len(src) / 1048576, len(raw) / 1048576))

    new_raw = process(raw, args.scale, args.colors)
    packed = zlib.compress(new_raw, 9)
    open(out_path, "wb").write(packed)

    print("输出文件: %s" % out_path)
    print("压缩后:   %.2f MB (protobuf %.2f MB)" % (len(packed) / 1048576, len(new_raw) / 1048576))
    print("体积变化: %.1f%%" % (100.0 * len(packed) / len(src)))


if __name__ == "__main__":
    main()
