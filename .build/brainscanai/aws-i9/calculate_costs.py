import csv
import json
from decimal import Decimal as D
from pathlib import Path

build=Path('.build/brainscanai/aws-i9')
rows=json.loads((build/'selected-prices.json').read_text(encoding='utf-8'))
row=next(r for r in rows if r['Instance Type']=='c7i.12xlarge')
assert row['vCPU']=='48' and row['Memory']=='96 GiB'
fx=D('1.1186')  # Same dated conversion as the user's source deck: ECB 8 October.
hour=(D(row['PricePerUnit'])+D('.005'))/fx
disk=D(50)*D('.0928')/fx
image_bytes=D(37272490)/D(1506)
storage_month=image_bytes*D(4000000)/D(2**30)*D('.024')/fx
requests=D(4000000)*(D('.0000053')+D('.00000042'))/fx
prototype_images=(D(37272490)/D(2**30)*D('.024')+D(1506)*(D('.0000053')+D('.00000042')))/fx
compute=hour*D('8.2')
values={
 'hour_eur':hour,'machine_30h':hour*30,'disk_month':disk,
 'prototype_images':prototype_images,'prototype_total':hour*30+disk+prototype_images,
 'compute_4m':compute,'storage_month':storage_month,'storage_year':storage_month*12,
 'storage_rolling':storage_month*D('6.5'),'disk_year':disk*12,'requests':requests,
 'one_shot':compute+storage_month*12+disk+requests,
 'rolling_year1':compute+storage_month*D('6.5')+disk*12+requests,
 'rolling_steady':compute+storage_month*12+disk*12+requests,
 'vm_24_7':hour*D(8760),
}
q=lambda value:value.quantize(D('.01'))
assert sum(map(q,[values['machine_30h'],disk,prototype_images]))==q(values['prototype_total'])
assert sum(map(q,[compute,storage_month*12,disk,requests]))==q(values['one_shot'])
assert sum(map(q,[compute,storage_month*D('6.5'),disk*12,requests]))==q(values['rolling_year1'])
assert sum(map(q,[compute,storage_month*12,disk*12,requests]))==q(values['rolling_steady'])
record={'fx_date':'2026-10-08','usd_per_eur':str(fx),'instance':row,
        'exact':{k:str(v) for k,v in values.items()},
        'display':{k:str(q(v)).replace('.',',') for k,v in values.items()}}
record['display']['hour_eur']=str(hour.quantize(D('.0001'))).replace('.',',')
(build/'costs-fixed-fx.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(record['display'],indent=2))
