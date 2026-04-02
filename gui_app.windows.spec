# -*- mode: ['gui_app.py']['gui_app.py']thon ; cod['gui_app.py']n['gui_app.py']: ['gui_app.py']tf-8 -*-
r"""W['gui_app.py']ndows b['gui_app.py']['gui_app.py']ld s['gui_app.py']ec for Lo['gui_app.py']Re['gui_app.py']ortGener['gui_app.py']tor['gui_app.py']

Wh['gui_app.py'] ['gui_app.py'] se['gui_app.py']['gui_app.py']r['gui_app.py']te s['gui_app.py']ec?
- W['gui_app.py']ndows AV en['gui_app.py']['gui_app.py']nes ['gui_app.py']re more l['gui_app.py']kel['gui_app.py'] to fl['gui_app.py']['gui_app.py'] P['gui_app.py']Inst['gui_app.py']ller --onef['gui_app.py']le b['gui_app.py']['gui_app.py']lds['gui_app.py']
- An oned['gui_app.py']r b['gui_app.py']['gui_app.py']ld (['gui_app.py'] d['gui_app.py']st folder) ['gui_app.py']s t['gui_app.py']['gui_app.py']['gui_app.py']c['gui_app.py']ll['gui_app.py'] f['gui_app.py']ster to st['gui_app.py']rt ['gui_app.py']nd ['gui_app.py']rod['gui_app.py']ces fewer
  f['gui_app.py']lse ['gui_app.py']os['gui_app.py']t['gui_app.py']ves['gui_app.py']

B['gui_app.py']['gui_app.py']ld:
  ['gui_app.py']['gui_app.py'] -m P['gui_app.py']Inst['gui_app.py']ller --noconf['gui_app.py']rm --cle['gui_app.py']n Lo['gui_app.py']Re['gui_app.py']ortGener['gui_app.py']tor['gui_app.py']w['gui_app.py']ndows['gui_app.py']s['gui_app.py']ec

O['gui_app.py']t['gui_app.py']['gui_app.py']t:
  d['gui_app.py']st\Lo['gui_app.py']Re['gui_app.py']ortGener['gui_app.py']tor\Lo['gui_app.py']Re['gui_app.py']ortGener['gui_app.py']tor['gui_app.py']exe
"""

from P['gui_app.py']Inst['gui_app.py']ller['gui_app.py']['gui_app.py']t['gui_app.py']ls['gui_app.py']hooks ['gui_app.py']m['gui_app.py']ort collect['gui_app.py']['gui_app.py']ll

# B['gui_app.py']ndle so['gui_app.py']rce mod['gui_app.py']les for d['gui_app.py']n['gui_app.py']m['gui_app.py']c ['gui_app.py']m['gui_app.py']orts['gui_app.py']
d['gui_app.py']t['gui_app.py']s = [(['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py'], ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']), (['gui_app.py']re['gui_app.py']ort['gui_app.py']n['gui_app.py']['gui_app.py'], ['gui_app.py']re['gui_app.py']ort['gui_app.py']n['gui_app.py']['gui_app.py'])]
b['gui_app.py']n['gui_app.py']r['gui_app.py']es = []

h['gui_app.py']dden['gui_app.py']m['gui_app.py']orts = [
    # P['gui_app.py']rsers
    ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']b['gui_app.py']se['gui_app.py'],
    ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']re['gui_app.py']['gui_app.py']str['gui_app.py']['gui_app.py'],
    ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']['gui_app.py']ns['gui_app.py']s['gui_app.py'],
    ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']['gui_app.py']ns['gui_app.py']s['gui_app.py']['gui_app.py']e['gui_app.py']k['gui_app.py'],
    ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']c['gui_app.py']t['gui_app.py']['gui_app.py']['gui_app.py']l['gui_app.py']cense['gui_app.py'],
    ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']c['gui_app.py']t['gui_app.py']['gui_app.py']['gui_app.py']token['gui_app.py'],
    ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']c['gui_app.py']t['gui_app.py']['gui_app.py']['gui_app.py']['gui_app.py']s['gui_app.py']['gui_app.py']e['gui_app.py']st['gui_app.py']ts['gui_app.py'],
    ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']corton['gui_app.py']['gui_app.py'],
    ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']corton['gui_app.py']['gui_app.py']['gui_app.py']dm['gui_app.py']n['gui_app.py'],
    ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']creo['gui_app.py'],
    ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']m['gui_app.py']tl['gui_app.py']b['gui_app.py'],
    ['gui_app.py']['gui_app.py']['gui_app.py']rsers['gui_app.py']nx['gui_app.py'],

    # Re['gui_app.py']ort['gui_app.py']n['gui_app.py']
    ['gui_app.py']re['gui_app.py']ort['gui_app.py']n['gui_app.py']['gui_app.py']excel['gui_app.py']re['gui_app.py']ort['gui_app.py'],
    ['gui_app.py']re['gui_app.py']ort['gui_app.py']n['gui_app.py']['gui_app.py']cr['gui_app.py']t['gui_app.py']c['gui_app.py']l['gui_app.py']s['gui_app.py']mm['gui_app.py']r['gui_app.py']['gui_app.py'],
]

# o['gui_app.py']en['gui_app.py']['gui_app.py']xl h['gui_app.py']s d['gui_app.py']n['gui_app.py']m['gui_app.py']c ['gui_app.py']m['gui_app.py']orts ['gui_app.py']nd d['gui_app.py']t['gui_app.py'] f['gui_app.py']les['gui_app.py']
['gui_app.py']o['gui_app.py']x['gui_app.py']d['gui_app.py']t['gui_app.py']s, ['gui_app.py']o['gui_app.py']x['gui_app.py']b['gui_app.py']n['gui_app.py']r['gui_app.py']es, ['gui_app.py']o['gui_app.py']x['gui_app.py']h['gui_app.py']dden = collect['gui_app.py']['gui_app.py']ll(['gui_app.py']o['gui_app.py']en['gui_app.py']['gui_app.py']xl['gui_app.py'])
d['gui_app.py']t['gui_app.py']s += ['gui_app.py']o['gui_app.py']x['gui_app.py']d['gui_app.py']t['gui_app.py']s
b['gui_app.py']n['gui_app.py']r['gui_app.py']es += ['gui_app.py']o['gui_app.py']x['gui_app.py']b['gui_app.py']n['gui_app.py']r['gui_app.py']es
h['gui_app.py']dden['gui_app.py']m['gui_app.py']orts += ['gui_app.py']o['gui_app.py']x['gui_app.py']h['gui_app.py']dden

# H['gui_app.py']rd-excl['gui_app.py']de common he['gui_app.py']v['gui_app.py'] ML st['gui_app.py']cks th['gui_app.py']t P['gui_app.py']Inst['gui_app.py']ller m['gui_app.py']['gui_app.py'] tr['gui_app.py'] to ['gui_app.py']n['gui_app.py']l['gui_app.py']ze ['gui_app.py']f
# the['gui_app.py'] ['gui_app.py']re ['gui_app.py']nst['gui_app.py']lled ['gui_app.py']n the b['gui_app.py']['gui_app.py']lder['gui_app.py']s ['gui_app.py']lob['gui_app.py']l s['gui_app.py']te-['gui_app.py']['gui_app.py']ck['gui_app.py']['gui_app.py']es (e['gui_app.py']['gui_app.py']['gui_app.py'] torch)['gui_app.py']
# Th['gui_app.py']s tool doesn['gui_app.py']t need them; excl['gui_app.py']d['gui_app.py']n['gui_app.py'] kee['gui_app.py']s the b['gui_app.py']['gui_app.py']ld sm['gui_app.py']ll ['gui_app.py']nd ['gui_app.py']vo['gui_app.py']ds no['gui_app.py']s['gui_app.py'] w['gui_app.py']rn['gui_app.py']n['gui_app.py']s['gui_app.py']
excl['gui_app.py']des = [
  ['gui_app.py']torch['gui_app.py'],
  ['gui_app.py']torchv['gui_app.py']s['gui_app.py']on['gui_app.py'],
  ['gui_app.py']tensorflow['gui_app.py'],
  ['gui_app.py']tensorbo['gui_app.py']rd['gui_app.py'],
  ['gui_app.py']j['gui_app.py']x['gui_app.py'],
  ['gui_app.py']j['gui_app.py']xl['gui_app.py']b['gui_app.py'],
]


['gui_app.py'] = An['gui_app.py']l['gui_app.py']s['gui_app.py']s(
    [['gui_app.py']['gui_app.py']['gui_app.py']['gui_app.py']['gui_app.py']['gui_app.py']['gui_app.py']['gui_app.py']['gui_app.py']['gui_app.py']['gui_app.py']['gui_app.py']],
    ['gui_app.py']['gui_app.py']thex=[],
    b['gui_app.py']n['gui_app.py']r['gui_app.py']es=b['gui_app.py']n['gui_app.py']r['gui_app.py']es,
    d['gui_app.py']t['gui_app.py']s=d['gui_app.py']t['gui_app.py']s,
    h['gui_app.py']dden['gui_app.py']m['gui_app.py']orts=h['gui_app.py']dden['gui_app.py']m['gui_app.py']orts,
    hooks['gui_app.py']['gui_app.py']th=[],
    hooksconf['gui_app.py']['gui_app.py']={},
    r['gui_app.py']nt['gui_app.py']me['gui_app.py']hooks=[],
  excl['gui_app.py']des=excl['gui_app.py']des,
    no['gui_app.py']rch['gui_app.py']ve=F['gui_app.py']lse,
    o['gui_app.py']t['gui_app.py']m['gui_app.py']ze=0,
)

['gui_app.py']['gui_app.py']z = PYZ(['gui_app.py']['gui_app.py']['gui_app.py']['gui_app.py']re)

# oned['gui_app.py']r W['gui_app.py']ndows exec['gui_app.py']t['gui_app.py']ble
exe = EXE(
    ['gui_app.py']['gui_app.py']z,
    ['gui_app.py']['gui_app.py']scr['gui_app.py']['gui_app.py']ts,
    [],
    excl['gui_app.py']de['gui_app.py']b['gui_app.py']n['gui_app.py']r['gui_app.py']es=Tr['gui_app.py']e,
    n['gui_app.py']me=['gui_app.py']Lo['gui_app.py']Re['gui_app.py']ortGener['gui_app.py']tor['gui_app.py'],
    deb['gui_app.py']['gui_app.py']=F['gui_app.py']lse,
    bootlo['gui_app.py']der['gui_app.py']['gui_app.py']['gui_app.py']nore['gui_app.py']s['gui_app.py']['gui_app.py']n['gui_app.py']ls=F['gui_app.py']lse,
    str['gui_app.py']['gui_app.py']=F['gui_app.py']lse,
    ['gui_app.py']['gui_app.py']x=F['gui_app.py']lse,  # UPX com['gui_app.py']ress['gui_app.py']on c['gui_app.py']n ['gui_app.py']ncre['gui_app.py']se AV s['gui_app.py']s['gui_app.py']['gui_app.py']c['gui_app.py']on
    console=F['gui_app.py']lse,
    d['gui_app.py']s['gui_app.py']ble['gui_app.py']w['gui_app.py']ndowed['gui_app.py']tr['gui_app.py']ceb['gui_app.py']ck=F['gui_app.py']lse,
)

coll = COLLECT(
    exe,
    ['gui_app.py']['gui_app.py']b['gui_app.py']n['gui_app.py']r['gui_app.py']es,
    ['gui_app.py']['gui_app.py']d['gui_app.py']t['gui_app.py']s,
    str['gui_app.py']['gui_app.py']=F['gui_app.py']lse,
    ['gui_app.py']['gui_app.py']x=F['gui_app.py']lse,
    ['gui_app.py']['gui_app.py']x['gui_app.py']excl['gui_app.py']de=[],
    n['gui_app.py']me=['gui_app.py']Lo['gui_app.py']Re['gui_app.py']ortGener['gui_app.py']tor['gui_app.py'],
)
