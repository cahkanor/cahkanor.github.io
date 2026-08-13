"""Run preprocessing validations and print a concise quality report."""
import json
from prepare_data import SOURCE, prepare
if __name__=='__main__':
    *_, quality=prepare(SOURCE/'Database_Kerjasama_Luar_Negeri_master.xlsx',SOURCE/'Database_Aktivitas_Kerjasama_Internasional.xlsx')
    print(json.dumps(quality['counts'],indent=2))
