"""Модуль конвертации Excel-файлов в JSON."""

import sys
import json

import pandas


class Conveter:
    """Конвертер из Excel в JSON."""

    def convert(self, data):
        """Конвертация."""
        side_effects = data.iloc[0].tolist()[1:]
        drugs = data.iloc[1:, 0].tolist()
        propobilities = data.iloc[1:, 1:].values.tolist()
        drug_side_effect_propobilities = {}
        for i, drug in enumerate(drugs):
            drug_side_effect_propobilities[drug] = []
            for j, effect in enumerate(side_effects):
                drug_side_effect_propobilities[drug].append(
                    (effect, propobilities[i][j]))
        return drug_side_effect_propobilities


def main():
    """Точка входа в программу."""
    if len(sys.argv) == 2:
        sheet_name = 'Common'
        data = pandas.read_excel(sys.argv[1],
                                 sheet_name=sheet_name,
                                 skiprows=1,
                                 usecols=lambda column: column != 0,
                                 header=None)
        converter = Conveter()
        with open('propobilities.json', 'w', encoding='utf-8') as f:
            json.dump(converter.convert(data), f, ensure_ascii=False, indent=4)

        print('Конвертация еспешно завершена')
    elif len(sys.argv) == 1:
        print('Ошибка! Не указано имя файла!')
    else:
        print('Ошибка! Указано слишком много аргументов!')


if __name__ == '__main__':
    main()
