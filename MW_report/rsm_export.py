# expected result: {link "A--B", Polarization, Channel_ID, Channel_BW, TxA_pwr, TxB_pwr, AntA_size, ANtA_height, AntB_size, AntB_height}

# >load antenna_list (name: , diameter:) to lookup diameter

# >load rsm
# - record_key=license_id

# > for record in rsm
#     > if licenseID not in record_key[]
#         > add license ID to record_key[]
#         > channelA_id=channel.replace('#','')
#         > polA=polarization
#         > linkname= Tx loc -- Rx loc # extract site codes with regex r"\d{3}-\d{3}", polarization,  channel_id, AntA_size=antenna_list['antenna_name'] or 'not found', AntA_height, TxA_pwr
#         > for record in rsm
#             > if licenseID not in record_key[] AND
#             > channelB_id=channel.replace('#','')
#             > polB=polarization
#             > if channelA_id=channelB_id AND polA=PolB
#                 > add license ID to record_key[]
#                 > grap AntB_size, AntB_height, TxB_pwr
#                 > break
# > write data to csv
#
# NOTES:
#   "Other" polarization - dual pol
#   "Non-specific" - in both cases related to Vertical, might need to recchek

import pandas as pd
import re
from pprint import pprint
import csv

ant_file = "C:/Work/01_Projects/0_architecture/MW_RFP/antennas.xlsx"
df = pd.read_excel(ant_file)
ant_records = df.to_dict(orient="records")
antennae = {}
for ant_ in ant_records:
    antennae[ant_["Tx Antenna Make"]] = ant_["Diameter"]


rsm_file = "C:/Work/01_Projects/0_architecture/MW_RFP/RMS/rsm.xlsx"
df = pd.read_excel(rsm_file)
rsm_records = df.to_dict(orient="records")

rsm_lic = []
link_export = []

replace_siteNames = {
    "32 MANUKA STREET TAUPO": "WKT-021-003-C",
    "SKY TOWER": "AKL-007-112-A",
}


def site_id(site_name):
    if site_name in replace_siteNames.keys():
        site_name = replace_siteNames[site_name]
    match_ = re.search(r"[A-Z]{3}-(\d{3}-\d{3})", str(site_name))
    if match_:
        site_id = match_.group(1)
    else:
        site_id = site_name
    return site_id


for record in rsm_records:
    lic_1 = str(record["Licence Id"])
    link_dict = {}
    if lic_1 not in rsm_lic:
        rsm_lic.append(lic_1)
        channelA_id = str(record["Channel"].replace("#", ""))

        polA_ = str(record["Polarisation"])
        if polA_ == "Other":
            polA_ = "Dual-pol"
        if polA_ == "Non-specific":
            polA_ = "Vertical"

        siteA1_name = str(record["Tx Loc Name"])
        siteB1_name = str(record["Rx Loc Name"])
        siteA_id = site_id(siteA1_name)
        siteB_id = site_id(siteB1_name)
        link_name = siteA_id + "--" + siteB_id
        link_dict["link_name"] = link_name
        link_dict["polarization"] = polA_
        link_dict["channel"] = channelA_id
        link_dict["emission"] = record["Emission"]
        link_dict["licenseA"] = lic_1
        link_dict["AntA_size"] = antennae[str(record["Tx Antenna Make"])]
        link_dict["AntA_hagl"] = record["Tx Antenna Height"]
        link_dict["SiteA_pwr"] = record["Power"]
        link_dict["Ref_frqA"] = record["Ref Freq"]

        for record in rsm_records:
            lic_2 = str(record["Licence Id"])
            if lic_2 not in rsm_lic:
                channelB_id = str(record["Channel"].replace("#", ""))

                polB_ = str(record["Polarisation"])
                if polB_ == "Other":
                    polB_ = "Dual-pol"
                if polB_ == "Non-specific":
                    polB_ = "Vertical"

                siteA2_name = str(record["Tx Loc Name"])
                siteB2_name = str(record["Rx Loc Name"])
                if (
                    siteA1_name == siteB2_name
                    and siteB1_name == siteA2_name
                    and channelA_id == channelB_id
                    and polA_ == polB_
                ):
                    rsm_lic.append(lic_2)
                    link_dict["licenseB"] = lic_2
                    link_dict["AntB_size"] = antennae[str(record["Tx Antenna Make"])]
                    link_dict["AntB_hagl"] = record["Tx Antenna Height"]
                    link_dict["SiteB_pwr"] = record["Power"]
                    link_dict["Ref_frqB"] = record["Ref Freq"]
                    break

        link_export.append(link_dict)

with open("link_export.csv", "w", encoding="utf8", newline="") as output_file:
    fc = csv.DictWriter(output_file, fieldnames=link_export[0].keys())
    fc.writeheader()
    fc.writerows(link_export)
