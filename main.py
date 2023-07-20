import config
import openpyxl
from tdotly import tly

COLUMNS = config.COLUMNS  # list of Excel column names(['A', 'B', 'C', ... 'AA', 'AB', 'AC'...])
FULL_TOKEN = "Bearer " + config.TOKEN  # Your API token
FILEPATH = config.FILEPATH  # Excel file path


def extract_url_id(url: str) -> str:
    return url.split('/')[-1]


def get_url_stats(url: str) -> dict:
    url_id = extract_url_id(url)
    short = tly.tly_shorturl()
    short.initialize(FULL_TOKEN)
    short.short_url_stats(short_url_id=url_id)
    return short.SHORT_URL_STATS


def get_urls_from_xls(filename: str) -> list:
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active
    urls = sheet['A']
    return [url.value for url in urls][1:]


def get_stats(urls: list) -> list:
    return [get_url_stats(url) for url in urls]


def get_dict_fields(stats: list) -> list:
    fields = set()
    for stat in stats:
        try:
            browsers = stat['browsers']
            for browser in browsers:
                fields.add('browser|' + browser['browser'] + '|total')
                fields.add('browser|' + browser['browser'] + '|unique_total')
            countries = stat['countries']
            for country in countries:
                fields.add('country|' + country['country_code'] + '|total')
                fields.add('country|' + country['country_code'] + '|unique_total')
            platforms = stat['platforms']
            for platform in platforms:
                fields.add('platform|' + platform['platform'] + '|total')
                fields.add('platform|' + platform['platform'] + '|unique_total')
            referrers = stat['referrers']
            for referrer in referrers:
                fields.add('referrer|' + referrer['referrer'] + '|total')
                fields.add('referrer|' + referrer['referrer'] + '|unique_total')
        except KeyError:
            continue
    fields = list(fields)
    fields.sort()
    return fields


def update_stats(stats: list) -> list:
    for stat_index in range(len(stats)):
        try:
            data = stats[stat_index]['data']
            stats[stat_index]['long_url'] = data['long_url']
            stats[stat_index]['created_at'] = data['created_at']
            stats[stat_index]['last_clicked'] = data['last_clicked']
            stats[stat_index]['total_clicks_last_thirty_days'] = data['total_clicks_last_thirty_days']
            stats[stat_index]['browsers'] = {
                browser['browser']: {'total': browser['total'], 'unique_total': browser['unique_total']} for browser in
                stats[stat_index]['browsers']}
            stats[stat_index]['countries'] = {
                browser['country_code']: {'total': browser['total'], 'unique_total': browser['unique_total']} for
                browser in stats[stat_index]['countries']}
            stats[stat_index]['platforms'] = {
                browser['platform']: {'total': browser['total'], 'unique_total': browser['unique_total']} for browser in
                stats[stat_index]['platforms']}
            stats[stat_index]['referrers'] = {
                browser['referrer']: {'total': browser['total'], 'unique_total': browser['unique_total']} for browser in
                stats[stat_index]['referrers']}
        except KeyError:
            pass
    return stats


if __name__ == '__main__':
    urls = get_urls_from_xls(FILEPATH)
    statistics = get_stats(urls)

    workbook = openpyxl.load_workbook(FILEPATH)
    sheet = workbook.active
    data_fields = ['long_url', 'clicks', 'unique_clicks', 'created_at', 'last_clicked', 'total_clicks_last_thirty_days',
                   'total_qr_scans']
    dict_fields = get_dict_fields(statistics)
    data_fields += dict_fields

    statistics = update_stats(statistics)

    for field_index in range(len(data_fields)):
        for index in range(len(urls) + 1):
            try:
                title = data_fields[field_index]
                if index == 0:
                    sheet[f'{COLUMNS[field_index + 1]}{index + 1}'] = title
                    continue

                if title.startswith('browser'):
                    browser_name = title.split('|')[-2]
                    stat_type = title.split('|')[-1]
                    if stat_type == 'unique_total': continue
                    browsers = statistics[index - 1]['browsers']

                    if browser_name in browsers.keys():
                        sheet[f'{COLUMNS[field_index + 1]}{index + 1}'] = browsers[browser_name]['total']
                        sheet[f'{COLUMNS[field_index + 2]}{index + 1}'] = browsers[browser_name]['unique_total']

                if title.startswith('country'):
                    country_name = title.split('|')[-2]
                    stat_type = title.split('|')[-1]
                    if stat_type == 'unique_total': continue
                    countries = statistics[index - 1]['countries']

                    if country_name in countries.keys():
                        sheet[f'{COLUMNS[field_index + 1]}{index + 1}'] = countries[country_name]['total']
                        sheet[f'{COLUMNS[field_index + 2]}{index + 1}'] = countries[country_name]['unique_total']

                if title.startswith('platform'):
                    platform_name = title.split('|')[-2]
                    stat_type = title.split('|')[-1]
                    if stat_type == 'unique_total': continue
                    platforms = statistics[index - 1]['platforms']

                    if platform_name in platforms.keys():
                        sheet[f'{COLUMNS[field_index + 1]}{index + 1}'] = platforms[platform_name]['total']
                        sheet[f'{COLUMNS[field_index + 2]}{index + 1}'] = platforms[platform_name]['unique_total']

                if title.startswith('referrer'):
                    referrer_name = title.split('|')[-2]
                    stat_type = title.split('|')[-1]
                    if stat_type == 'unique_total': continue
                    referrer = statistics[index - 1]['referrers']

                    if referrer_name in referrer.keys():
                        sheet[f'{COLUMNS[field_index + 1]}{index + 1}'] = referrer[referrer_name]['total']
                        sheet[f'{COLUMNS[field_index + 2]}{index + 1}'] = referrer[referrer_name]['unique_total']

                sheet[f'{COLUMNS[field_index + 1]}{index + 1}'] = statistics[index - 1][data_fields[field_index]]
            except KeyError:
                pass

    workbook.save(FILEPATH)
