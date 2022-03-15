from . import util


def handle_breakdown_accounts(rpes, kwargs):
    shape_id_to_text = kwargs.get('shape_id_to_text', {})
    accounts_to_be_broken_down = kwargs.get('breakdown_accounts', ())
    operator_accounts = kwargs.get('operators', ())
    rpe_dict = util.create_rpe_dictionary(rpes)

    breakdown_account_dictionary = {}
    new_rpe_dictionary = {}
    for account_to_be_broken_down in accounts_to_be_broken_down:
        rpe_of_breakdown_account = rpe_dict.get(account_to_be_broken_down, ())
        if len(rpe_of_breakdown_account) > 2:
            operator = rpe_of_breakdown_account[2]
            new_rpe = []
            breakdown_account_id = 0
            for n, element in enumerate(rpe_of_breakdown_account):
                if element not in operator_accounts:
                    # Breakdown Account Dictionary Creation
                    new_breakdown_account = f'breakdown_of_account_{account_to_be_broken_down}_{breakdown_account_id}'
                    breakdown_account_id += 1
                    if account_to_be_broken_down in breakdown_account_dictionary:
                        breakdown_account_dictionary[account_to_be_broken_down].append(new_breakdown_account)
                    else:
                        breakdown_account_dictionary[account_to_be_broken_down] = [new_breakdown_account]

                    # Adding New Direct Links and ShapeID to Text to kwargs
                    from_id = element
                    to_id = new_breakdown_account
                    new_direct_link = (from_id, to_id, 0)
                    kwargs['direct_links'] += (new_direct_link,)
                    kwargs['shape_id_to_text'][new_breakdown_account] = shape_id_to_text.get(element)

                    # New RPE Creation
                    new_rpe.append(new_breakdown_account)
                    if n > 0:
                        new_rpe.append(operator)
            new_rpe_dictionary[account_to_be_broken_down] = (account_to_be_broken_down, tuple(new_rpe))

    # Updating RPEs: Account now adds its breakdowns.
    new_rpes = []
    for each_rpes in rpes:
        account = each_rpes[0]
        if account in new_rpe_dictionary:
            new_rpe = new_rpe_dictionary[account]
            new_rpes.append(new_rpe)
        else:
            new_rpes.append(each_rpes)

    # Inserting breakdown accounts
    worksheets_data = kwargs.get('sheets_data', {})
    for sheet_name, sheet_contents in worksheets_data.items():
        new_sheet_contents = []
        for n, content in enumerate(sheet_contents):
            if content in breakdown_account_dictionary:
                account_to_be_broken_down = content

                if n > 0:
                    if str(new_sheet_contents[n - 1]) != 'blank':
                        new_sheet_contents.append('blank')
                new_sheet_contents += breakdown_account_dictionary[account_to_be_broken_down]
                new_sheet_contents.append(content)
                new_sheet_contents.append('blank')
            else:
                new_sheet_contents.append(content)

        # Modifying kwargs itself, rather than returning new value
        kwargs['sheets_data'][sheet_name] = new_sheet_contents
    return tuple(new_rpes)
