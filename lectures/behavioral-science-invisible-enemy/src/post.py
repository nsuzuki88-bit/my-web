# pptxgenjs 出力の後処理：東アジアフォントを BIZ UDゴシック（日本語charset）に統一し、テーマフォントも設定
import sys, zipfile, shutil, re
src, dst = sys.argv[1], sys.argv[2]
FONT = "BIZ UDゴシック"
zin = zipfile.ZipFile(src)
zout = zipfile.ZipFile(dst + ".tmp", "w", zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename.endswith(".xml"):
        s = data.decode("utf-8")
        s = s.replace(f'<a:ea typeface="{FONT}" pitchFamily="34" charset="-122"/>',
                      f'<a:ea typeface="{FONT}" pitchFamily="49" charset="-128"/>')
        if item.filename.startswith("ppt/theme/"):
            s = re.sub(r'<a:latin typeface="[^"]*"( panose="[^"]*")?/>', f'<a:latin typeface="{FONT}"/>', s)
            s = s.replace('<a:ea typeface=""/>', f'<a:ea typeface="{FONT}"/>')
            s = re.sub(r'<a:font script="Jpan" typeface="[^"]*"/>', f'<a:font script="Jpan" typeface="{FONT}"/>', s)
        if item.filename.startswith("ppt/charts/chart"):
            s = s.replace('typeface="Arial"', f'typeface="{FONT}"')
            s = re.sub(r'sz="1\d{3}"', 'sz="2000"', s)
            s = re.sub(rf'<a:latin typeface="{FONT}"\s*/>(?!\s*<a:ea)', f'<a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/>', s)
        data = s.encode("utf-8")
    zout.writestr(item, data)
zout.close()
shutil.move(dst + ".tmp", dst)
print("post-processed ->", dst)
