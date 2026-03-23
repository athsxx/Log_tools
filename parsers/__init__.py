from .catia_license import parse_files as parse_catia_license
from .catia_token import parse_files as parse_catia_token
from .catia_usage_stats import parse_files as parse_catia_usage_stats
from .ansys import parse_files as parse_ansys
from .ansys_peak import parse_files as parse_ansys_peak
from .cortona import parse_files as parse_cortona
from .cortona_admin import parse_files as parse_cortona_admin
from .nx import parse_files as parse_nx
from .creo import parse_files as parse_creo
from .matlab import parse_files as parse_matlab

PARSER_MAP = {
    "catia_license": parse_catia_license,
    "catia_token": parse_catia_token,
    "catia_usage_stats": parse_catia_usage_stats,
    "ansys": parse_ansys,
    "ansys_peak": parse_ansys_peak,
    "cortona": parse_cortona,
    "cortona_admin": parse_cortona_admin,
    "nx": parse_nx,
    "creo": parse_creo,
    "matlab": parse_matlab,
}
