import rpe_to_normal


def set_formulas(nop, rpes_converted: tuple):
    op = rpe_to_normal.operator
    for account, rpe in rpes_converted:
        formulas_list = []
        for n in range(nop):
            rpe_n_list = []
            for element in rpe:
                if element in op.values():
                    operator = element
                    rpe_n_list.append(operator)
                else:
                    rpe_n_list.append(element[n])
            rpe_n = tuple(rpe_n_list)
            formula = rpe_to_normal.post_to_infix(rpe_n)
            formulas_list.append(f'={formula}')
        formulas = tuple(formulas_list)
        account.set_values(formulas)
