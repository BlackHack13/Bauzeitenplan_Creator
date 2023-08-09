import calendar
import datetime

def spalten(startjahr, startmonat, endjahr, endmonat):
    jahre = [x for x in range(startjahr, endjahr + 1)]
    monate = [x for x in range(1, 12 + 1)]

    kombi = [(jahr, monat) for jahr in jahre for monat in monate]
    kombi = kombi[startmonat - 1:endmonat - 12]
    kombi_mit_kalenderwochen = []

    for jahr, monat in kombi:
        kalenderwochen = set()
        for x in calendar.Calendar().itermonthdays4(jahr, monat):
            kalenderwochen.add(datetime.date(x[0], x[1], x[2]).isocalendar()[1])
        kalenderwochen = sorted(list(kalenderwochen))
        for kalenderwoche in kalenderwochen:
            kombi_mit_kalenderwochen.append((jahr, monat, kalenderwoche))
    return kombi_mit_kalenderwochen

result = spalten(2022, 8, 2024, 1)

for x in result:
    print(x)
