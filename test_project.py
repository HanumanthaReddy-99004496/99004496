"""
test_project.py
author: Hanumantha Reddy
PS.no: 99004496
"""

import project

obj = project.Excel()  # creating object to call class methods


def test_psno():
    """
    view all 15 ps numbers existed in excel file
    """
    assert project.view_psnumbers() == ['PS number', 99004480, 99004481,
                                        99004482, 99004483,
                                        99004484, 99004485,
                                        99004486, 99004487,
                                        99004488, 99004489,
                                        99004490, 99004491,
                                        99004492, 99004493,
                                        99004494]


def test_validate_input_ps():
    """
    validate if input ps.no is existed or not
    """
    assert obj.validate(99004496) == "Invalid PS.No"
    assert obj.validate(99004481) == "Valid"


def test_view_sheets():
    """view sheets"""
    assert project.all_sheets == ['academic', 'tests', 'calories', 'cuisine', 'diet']


def test_validate_sheet():
    """test validate sheet names in xl file"""
    assert obj.validate_sheet('diet') == "Valid"
    assert obj.validate_sheet(('abcd')) == "Invalid"


def test_get_data():
    """test user requested data"""
    assert obj.get_data(99004491, 'calories') == [99004491, 720, 4, 420, 2, 'chocolate,'
                                                  ' pasta, soup,' ' chips, popcorn',
                                                  'sadness, stress, cold weather', 3, 3, 3, 1,
                                                  'I am very health concious. I eat many'
                                                  ' fruits, veggies, '
                                                  'and protiens. ',
                                                  1, 1, 'Less meat. ', 4, 5, 1, 2, 5]


def test_write():
    """test for writing into output.xlsx"""
    assert project.writeto_outputxl() == "Written output to 'output.xlsx'"
