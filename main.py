import sys
import functions

tool_title = """\n
                   ---
                ---------
        --------------------------    
 ERP  data  retrieve  tool  from  JD Edwards
        --------------------------    
                ---------
                   ---
                   
"""

options = """-----

Please choose option you would like to run this time or press 0 for EXIT:

1. Login into JDE
2. Export control - data retrieve from JDE
3. Unit cost - data retrieve from JDE
4. Branch code change
0. EXIT
"""

# branch codes in JDE
branch1_c = 'branch1_code_number_here'
branch2_c = 'branch2_code_number_here'

def exit_program():
    print("Goodbye !")
    sys.exit()

while True:
    username = functions.username_choice
    password = functions.password_session
    if username == '0' or username == 'esc':
        exit_program()
    elif username != '0':
        print(tool_title,"\n-----")
        print(f"Branch Code 1: |{branch1_c}|")
        print(f"Branch Code 2: |{branch2_c}|")
        user_choice = input(options)
        if user_choice == '1':
            functions.run_login_to_jde()
            break
        elif user_choice == '2':
            functions.define_range_of_rows()
            functions.run_export_control_script()
        elif user_choice == '3':
            functions.define_range_of_rows()
            functions.run_unit_cost_script()
        elif user_choice == '4':
            branch1_c, branch2_c = functions.branch_code_change()
        elif user_choice == '0' or username == 'esc':
            exit_program()
