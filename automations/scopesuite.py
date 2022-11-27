from evsauto.webdriver.chrome import CustomChromeWebDriver


def submit_grades(webdriver_wrapper: CustomChromeWebDriver, username: str, password: str) -> None:
    """
    Submits the highest possible grade for today in ScopeSuite.
    """

    # open scopesuite
    webdriver_wrapper.get(
        "https://amzmanagement-us.herokuapp.com/amz/index.html#/ojl")

    # find all necessary login elements
    username_input = webdriver_wrapper.find_element(
        '//div[@class="login-content"]/input[@type="text"]')
    password_input = webdriver_wrapper.find_element(
        '//div[@class="login-content"]/input[@type="password"]')
    submit_button = webdriver_wrapper.find_element(
        '//div[@class="login-content"]/button')

    # login to scopesuite
    webdriver_wrapper.type_into_element(username_input, username)
    webdriver_wrapper.type_into_element(password_input, password)
    webdriver_wrapper.click_element(submit_button)

    # find all necessary form elements to submit today's grade
    score_inputs = webdriver_wrapper.find_elements(
        '//input[@aria-valuemin="0" and @aria-valuemax="5"]')
    score_strengths = webdriver_wrapper.find_elements(
        '//div[@class="attitude-item positive"]')
    submit_button = webdriver_wrapper.find_element(
        '//button/span[contains(text(), "Submit")]/..')

    # loop through score inputs and give myself the highest score
    for score_input in score_inputs:
        webdriver_wrapper.type_into_element(score_input, '5')

    # loop through strenghts/weaknesses and give myself all strengths
    for score_strength in score_strengths:
        webdriver_wrapper.click_element(score_strength)

    # submit grade for today
    webdriver_wrapper.click_element(submit_button)

    # ADD CLICKING LOGOUT
