import xlrd
import pathlib


class Component:
    def __init__(self, number: int, name: str, PU, TA, TC, TE, TU, NVX, B, NVP, RB, T_k, p_k, pv):
        self.number = number
        self.name = name
        self.PU = PU  # Вероятность того, что компонент будет использоваться
        # self.PF = PF  # Вероятность того, что в компоненте возникнет сбой
        self.PF = 1 - R(p_k, pv)
        self.TA = TA  # Относительное время доступа к компоненту
        self.TC = TC  # Относительное время анализа сбоя в компоненте
        self.TE = TE  # Относительное время устранения сбоя в компоненте
        self.TU = TU  # Относительное время использования компонента
        # self.T = T  # Трудоемкость разработки компонента
        self.T_k = T_k  # Трудоемкость разработки версии k компонента в чел-часах
        self.p_k = p_k  # Вероятность безотказной работы версии
        self.pv = pv  # Вероятность безотказной работы алгоритма голосования
        self.NVX = NVX  # Трудоемкость разработки среды исполнения версий
        # (приемочного теста для RB (recovery block, блок восстановления)
        # или алгоритма голосования для NVP (N-version programming, N-версионное программирование)
        self.B = B  # Бинарная переменная, принимающая значение 1 (тогда NVP=0, RB=0),
        # если в программном компоненте не используется программная избыточность, иначе равна 0
        self.NVP = NVP  # Бинарная переменная, принимающая значение 1 (тогда B=0, RB=0),
        # если в программном компоненте введена программная избыточность методом NVP, иначе равна 0
        self.RB = RB  # Бинарная переменная, принимающая значение 1 (тогда B=0, NVP=0),
        # если в программном компоненте введена программная избыточность методом RB, иначе равна 0.

    def __str__(self):
        Tk_string = ''.join('{:<9}'.format(str(e)) for e in self.T_k)
        pk_string = ''.join('{:<9}'.format(str(e)) for e in self.p_k)

        return '{:<5} {:<10} {:<5} {:<5} {:<5} {:<5} {:<5} {:<7} {:<7} {:<3} {:<3} {:<3} {:<15} {:<15}'.format(
            self.number, self.name, self.PU, round(self.PF, 2), self.TA, self.TC, self.TE, self.TU, self.NVX, self.B, self.NVP, self.RB, Tk_string, pk_string)


def architecture_to_str(architecture: list[list[Component]]):
    result = '{:<5} {:<10} {:<5} {:<5} {:<5} {:<5} {:<5} {:<7} {:<7} {:<3} {:<3} {:<3} {:<7} {:<27} {:<35}\n'.format(
            '№', 'Название', 'PU', 'PF', 'TA', 'TC', 'TE', 'TU', 'NVX', 'B', 'NVP', 'RB', 'T', 'Tk', 'pk')
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


def get_from_excel(path):
    wb = xlrd.open_workbook(path)
    params = wb.sheet_by_index(0)
    deps = wb.sheet_by_index(1)
    vers = wb.sheet_by_index(2)

    dependencies = []
    total_levels = int(max(params.col_values(1)[1:]))  # Определить общее количество уровней
    total_components = int(params.col_values(0)[-1])  # Определить общее количество компонент
    architecture = [[] for i in range(total_levels)]  # Создать архитектуру

    for row in range(1, params.nrows):  # Пропустить шапку таблицы
        values = params.row_values(row)  # Считать значения ячеек в текущей строке
        # Проименовать значения для удобства
        number, level, name, PU, TA, TC, TE, TU, NVX, B, NVP, RB, pv =\
            int(values[0]), int(values[1]), values[3], values[4], values[5], values[6], values[7], values[8], values[9], int(values[10]), int(values[11]), int(values[12]), values[13]

        row_end_index = -1  # Индекс, на котором заканчиваются входные данные для текущей компоненты
        for cell in vers.row_values(row):
            if cell == '':
                break
            else:
                row_end_index += 1

        T_k = vers.row_values(row)[1:row_end_index:2]
        p_k = vers.row_values(row)[2:row_end_index+1:2]

        architecture[level - 1].append(Component(number, name, PU, TA, TC, TE, TU, NVX, B, NVP, RB, T_k, p_k, pv))  # Добавить компоненту на уровень
        dependencies.append(deps.row_values(row)[1:])  # Добавить все компоненты, которые зависят от текущей
    return architecture, dependencies


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
            # print(component.number, 'зависит от', index + 1, dependencies[index][component.number - 1])
            result.append(i)

    return result


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
                        result_ml += dependencies[mn.number - 1][ml.number - 1] * (ml.TA + ml.TC + ml.TE)
                    result_mn_1 += dependencies[ji.number - 1][mn.number - 1] * ((mn.TA + mn.TC + mn.TE) + result_ml)

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
                            result_ml += dependencies[mn.number - 1][ml.number - 1] * (ml.TA + ml.TC + ml.TE)
                        result_mn_2 += dependencies[jk.number - 1][mn.number - 1] * ((mn.TA + mn.TC + mn.TE) + result_ml)
                result_jk += dependencies[ji.number - 1][jk.number - 1] * ((jk.TA + jk.TC + jk.TE) + result_mn_2)

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
                        result_ml += (1 - dependencies[mn.number - 1][ml.number - 1]) * ml.TU
                    result_mn_1 += (1 - dependencies[ji.number - 1][mn.number - 1]) * (mn.TU + result_ml)

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
                            result_ml += (1 - dependencies[mn.number - 1][ml.number - 1]) * ml.TU
                        result_mn_2 += (1 - dependencies[jk.number - 1][mn.number - 1]) * (mn.TU + result_ml)
                result_jk += (1 - dependencies[ji.number - 1][jk.number - 1]) * (jk.TU + result_mn_2)

            result_ji += ji.PU * (1 - ji.PF) * (ji.TU + result_mn_1 + result_jk)
    return result_ji


def S():  # Функция готовности системы
    mttf = MTTF(architecture, dependencies)
    tr = TR(architecture, dependencies)
    return mttf / (tr + mttf)


def T_s(architecture: list[list[Component]]):  # Общая трудоемкость реализации программной системы
    result = 0
    for j in range(0, len(architecture)):
        for i in range(0, len(architecture[j])):
            a = architecture[j][i]

            sum_T = 0
            for t in a.T_k[1:]:
                sum_T += t

            result += a.B * a.T_k[0] + (a.NVP + a.RB) * (a.NVX + sum_T)
    return result


def R(p_k, pv):  # Надёжность мультиверсионного компонента
    result = 1
    for pk in p_k:
        result *= (1 - pk)

    return pv * (1 - result)


if __name__ == '__main__':
    architecture, dependencies = get_from_excel(pathlib.Path(__file__).parent / 'data_modified.xlsx')
    print(architecture_to_str(architecture))
    print(dependencies_to_str(dependencies))
    print('Среднее время простоя', TR(architecture, dependencies))
    print('Среднее время сбоя', MTTF(architecture, dependencies))
    print('Коэффициент готовности', S())
    print('Трудозатраты', T_s(architecture))

    print('\nВероятность безотказной работы каждого компонента')
    n = 1
    for l, level in enumerate(architecture):
        for c, component in enumerate(level):
            print(f'{n}\t[{l+1}][{c+1}] \t {1 - component.PF}')
            n += 1
