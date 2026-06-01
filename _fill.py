# -*- coding: utf-8 -*-
import nbformat
path = r'c:\Users\a_eug\Eugene\Kuliah\ProbStat\ALP-ProbStat\Final.ipynb'
nb = nbformat.read(path, as_version=4)
hits = [c for c in nb.cells if 'disamakan jadi ...' in c.source]
assert len(hits) == 1, f'got {len(hits)} matches'
hits[0].source = hits[0].source.replace('disamakan jadi ...', 'disamakan jadi `SUV SMALL`')
nbformat.validate(nb)
nbformat.write(nb, path)
print('filled blank ->', '`SUV SMALL`')
