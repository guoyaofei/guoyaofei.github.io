from pathlib import Path

p = Path(__file__).with_name('build_core10.py')
src = p.read_text(encoding='utf-8')
marker = "if OUT.exists(): shutil.rmtree(OUT)"
if marker in src:
    src = src.split(marker, 1)[0]
append = r'''
archive_base = OUT / "开江中学实验学校红专学校发展规划首批十份核心文件"
shutil.make_archive(str(archive_base), "zip", OUT, ROOT.name)
print("Generated Word files:", len(list(ROOT.glob("*.docx"))))
print(str(archive_base) + ".zip")
'''
exec(compile(src + append, str(p), 'exec'), {'__name__': '__main__'})
