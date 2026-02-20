"""One-time script to process raw FBI UCR crime data into a clean lookup CSV."""
import pandas as pd
import re

STATE_ABBR = {
    'ala': 'AL', 'alaska': 'AK', 'ariz': 'AZ', 'ark': 'AR', 'calif': 'CA',
    'colo': 'CO', 'conn': 'CT', 'del': 'DE', 'fla': 'FL', 'ga': 'GA',
    'hawaii': 'HI', 'idaho': 'ID', 'ill': 'IL', 'ind': 'IN', 'iowa': 'IA',
    'kan': 'KS', 'ky': 'KY', 'la': 'LA', 'maine': 'ME', 'md': 'MD',
    'mass': 'MA', 'mich': 'MI', 'minn': 'MN', 'miss': 'MS', 'mo': 'MO',
    'mont': 'MT', 'neb': 'NE', 'nev': 'NV', 'nh': 'NH', 'nj': 'NJ',
    'nm': 'NM', 'ny': 'NY', 'nc': 'NC', 'nd': 'ND', 'ohio': 'OH',
    'okla': 'OK', 'ore': 'OR', 'pa': 'PA', 'ri': 'RI', 'sc': 'SC',
    'sd': 'SD', 'tenn': 'TN', 'texas': 'TX', 'utah': 'UT', 'vt': 'VT',
    'va': 'VA', 'wash': 'WA', 'wva': 'WV', 'wis': 'WI', 'wyo': 'WY',
    'dc': 'DC',
}


def parse_city_state(name):
    name = name.strip()
    if 'national' in name.lower():
        return 'National', 'US'
    parts = name.rsplit(',', 1)
    city = parts[0].strip()
    if len(parts) > 1:
        state_raw = parts[1].strip().lower()
        state_key = state_raw.replace('.', '').replace(' ', '')
        state = STATE_ABBR.get(state_key, state_raw.upper().replace('.', ''))
    else:
        state = ''
    # Remove county/parish suffix
    city = re.sub(r'\s+(county|parish)$', '', city, flags=re.IGNORECASE).strip()
    return city, state


df = pd.read_csv('data/crime_data_raw.csv')
df2015 = df[df['year'] == 2015].copy()

rows = []
for _, row in df2015.iterrows():
    city, state = parse_city_state(row['department_name'])
    if city == 'National':
        continue
    rows.append({
        'city': city,
        'state': state,
        'original_name': row['department_name'],
        'violent_per_100k': round(row['violent_per_100k'], 1) if pd.notna(row['violent_per_100k']) else None,
        'homicide_per_100k': round(row['homs_per_100k'], 2) if pd.notna(row['homs_per_100k']) else None,
        'robbery_per_100k': round(row['rob_per_100k'], 1) if pd.notna(row['rob_per_100k']) else None,
        'assault_per_100k': round(row['agg_ass_per_100k'], 1) if pd.notna(row['agg_ass_per_100k']) else None,
        'data_year': 2015,
    })

out = pd.DataFrame(rows)
out.to_csv('data/crime_data.csv', index=False)
print(f'Saved {len(out)} cities to data/crime_data.csv')
print(out[['city', 'state', 'violent_per_100k']].to_string())
