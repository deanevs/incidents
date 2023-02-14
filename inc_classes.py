from dataclasses import dataclass
from datetime import datetime

@dataclass
class Incident:
    int_ref: str
    safety_not_ref: str
    inc_date: datetime
    rec_date: datetime
    added_date: datetime
    action_date: datetime
    due_date: datetime
    closed_date: datetime
    alert_form: str
    inc_desc: str
    clin_conseq: str
    actions_taken: str
    declared_by: str
    issued_by: str

    def delta_days(self):
        try:
            delta = (self.inc_date - self.closed_date).days
            return ('Total days {}'.format(delta))
        except:
            return 'N/A'

@dataclass
class Asset:
    asset_id: str
    equip_no: str
    gmdn: str
    manu: str
    model: str
    serial: str

