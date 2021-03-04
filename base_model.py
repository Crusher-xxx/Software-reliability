import xlrd
import pathlib
import math


class Component:
    def __init__(self, number: int, name: str, PU, PF, TA, TC, TE, TU):
        self.number = number
        self.name = name
        self.PU = PU  # Вероятность того, что компонент будет использоваться
        self.PF = PF  # Вероятность того, что в компоненте возникнет сбой
        self.TA = TA  # Относительное время доступа к компоненту
        self.TC = TC  # Относительное время анализа сбоя в компоненте
        self.TE = TE  # Относительное время устранения сбоя в компоненте
        self.TU = TU  # Относительное время использования компонента

    def __str__(self):
        return '{:<5} {:<10} {:<5} {:<5} {:<5} {:<5} {:<5} {:<5}'.format(
            self.number, self.name, self.PU, self.PF, self.TA, self.TC, self.TE, self.TU)


def architecture_to_str(architecture: list[list[Component]]):
    result = '{:<5} {:<10} {:<5} {:<5} {:<5} {:<5} {:<5} {:<5}\n'.format(
            '№', 'Название', 'PU', 'PF', 'TA', 'TC', 'TE', 'TU')
    for level in architecture:
        for component in level:
            result += str(component) + '\n'
    return result


def dependencies_to_str(dependencies):
    table = ''

    for horizontal_counter in range(0, len(dependencies[0]) + 1):
        table += '{:<10}'.format(horizontal_counter)
    table += '\n'

    vertical_counter = 1
    for level in dependencies:
        table += '{:<10}'.format(vertical_counter)
        for comp in level:
            table += '{:<10}'.format(comp)
        table += '\n'
        vertical_counter += 1

    return table


def get_dependent_indices(dependencies, architecture: list[list[Component]], at_level, on_component):
    # Индексы элементов, зависящих от элемента[at_level][on_component] на уровне at_level
    result = []
    # Индекс компоненты, от которой ищутся зависимые
    index = architecture[at_level][on_component].number - 1

    # Компоненты, среди которых вести поиск (общий уровень с предыдущей компонентой)
    for i, component in enumerate(architecture[at_level]):
        if on_component == i:  # Проверка на сравнение с собой
            continue
        if dependencies[index][component.number - 1] != 0:  # Если есть зависимость
            result.append(i)

    return result


def get_from_excel(path):
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    dependencies = []
    levels = int(max(sheet.col_values(1)[1:]))  # Найти количество уровней
    components = [[] for i in range(levels)]
    for row in range(1, sheet.nrows):  # Пропустить шапку таблицы
        values = sheet.row_values(row)  # Считать значения ячеек в текущей строке
        # Проименовать значения для удобства
        number, level, name, PU, PF, TA, TC, TE, TU = \
            int(values[0]), int(values[1]), values[3], values[4], values[5], values[6], values[7], values[8], values[9]

        components[level - 1].append(Component(number, name, PU, PF, TA, TC, TE, TU))  # Добавить компоненту на уровень
        dependencies.append(values[10:])  # Добавить все компоненты, которые зависят от текущей
    return components, dependencies


def TR(architecture: list[list[Component]], dependencies):  # Среднее время простоя всей системы
    result_ji = 0
    for j in range(len(architecture)):  # Количество уровней в архитектуре
        for i in range(len(architecture[j])):  # Количество компонент на данном уровне
            ji = architecture[j][i]

            result_mn_1 = 0
            for m in range(len(architecture)):
                if m == j:
                    continue
                for n in range(len(architecture[m])):
                    mn = architecture[m][n]

                    result_ml = 0
                    for l in get_dependent_indices(dependencies, architecture, m, n):
                        ml = architecture[m][l]
                        result_ml += dependencies[ml.number - 1][mn.number - 1] * (ml.TA + ml.TC + ml.TE)
                    result_mn_1 += dependencies[mn.number - 1][ji.number - 1] * ((mn.TA + mn.TC + mn.TE) + result_ml)

            result_jk = 0
            for k in get_dependent_indices(dependencies, architecture, j, i):
                jk = architecture[j][k]

                result_mn_2 = 0
                for m in range(len(architecture)):
                    if m == j:
                        continue
                    for n in range(len(architecture[m])):
                        mn = architecture[m][n]

                        result_ml = 0
                        for l in get_dependent_indices(dependencies, architecture, m, n):
                            ml = architecture[m][l]
                            result_ml += dependencies[ml.number - 1][mn.number - 1] * (ml.TA + ml.TC + ml.TE)
                        result_mn_2 += dependencies[mn.number - 1][jk.number - 1] * ((mn.TA + mn.TC + mn.TE) + result_ml)
                result_jk += dependencies[jk.number - 1][ji.number - 1] * ((jk.TA + jk.TC + jk.TE) + result_mn_2)

            result_ji += ji.PU * ji.PF * ((ji.TA + ji.TC + ji.TE) + result_mn_1 * result_jk)

    return result_ji


def MTTF(architecture: list[list[Component]], dependencies):  # Среднее время появления сбоя во всей системе
    result_ji = 0
    for j in range(len(architecture)):  # Количество уровней в архитектуре
        for i in range(len(architecture[j])):  # Количество компонент на данном уровне
            ji = architecture[j][i]

            result_mn_1 = 0
            for m in range(len(architecture)):
                if m == j:
                    continue
                for n in range(len(architecture[m])):
                    mn = architecture[m][n]

                    result_ml = 0
                    for l in get_dependent_indices(dependencies, architecture, m, n):
                        ml = architecture[m][l]
                        result_ml += (1 - dependencies[ml.number - 1][mn.number - 1]) * ml.TU
                    result_mn_1 += (1 - dependencies[mn.number - 1][ji.number - 1]) * (mn.TU + result_ml)

            result_jk = 0
            for k in get_dependent_indices(dependencies, architecture, j, i):
                jk = architecture[j][k]

                result_mn_2 = 0
                for m in range(len(architecture)):
                    if m == j:
                        continue
                    for n in range(len(architecture[m])):
                        mn = architecture[m][n]

                        result_ml = 0
                        for l in get_dependent_indices(dependencies, architecture, m, n):
                            ml = architecture[m][l]
                            result_ml += (1 - dependencies[ml.number - 1][mn.number - 1]) * ml.TU
                        result_mn_2 += (1 - dependencies[mn.number - 1][jk.number - 1]) * (mn.TU + result_ml)
                result_jk += (1 - dependencies[jk.number - 1][ji.number - 1]) * (jk.TU + result_mn_2)

            result_ji += ji.PU * (1 - ji.PF) * (ji.TU + result_mn_1 + result_jk)
    return result_ji


def S(t):  # Функция готовности системы
    i = 2  # Состояние системы 0 - отказ, 1 - функционирует
    mttf = MTTF(architecture, dependencies)
    tr = TR(architecture, dependencies)
    a = mttf / (tr + mttf)
    b = -((tr + mttf) / (tr * mttf)) * t

    if i == 0:
        return a - a * (math.e ** b)
    elif i == 1:
        return a + a * (math.e ** b)
    elif i == 2:
        return a


if __name__ == '__main__':
    architecture, dependencies = get_from_excel(pathlib.Path(__file__).parent / 'data_base.xlsx')
    print(architecture_to_str(architecture))
    print(dependencies_to_str(dependencies))
    print(TR(architecture, dependencies))
    print(MTTF(architecture, dependencies))
    print(S(3))
