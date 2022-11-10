from custom_webdriver import CustomChromeWebDriver
import os


def submit_grades(chrome: CustomChromeWebDriver):
    """
    Submits the highest possible grade for today in ScopeSuite.
    """

    # grab login details from environment vars
    SCOPESUITE_USERNAME = os.getenv("SCOPESUITE_USERNAME")
    SCOPESUITE_PASSWORD = os.getenv("SCOPESUITE_PASSWORD")

    # open scopesuite
    chrome.get("https://amzmanagement-us.herokuapp.com/amz/index.html#/ojl")

    # find all necessary login elements
    username_input = chrome.find_element(
        '//div[@class="login-content"]/input[@type="text"]')
    password_input = chrome.find_element(
        '//div[@class="login-content"]/input[@type="password"]')
    submit_button = chrome.find_element(
        '//div[@class="login-content"]/button')

    # login to scopesuite
    chrome.type_into_element(username_input, SCOPESUITE_USERNAME)
    chrome.type_into_element(password_input, SCOPESUITE_PASSWORD)
    chrome.click_element(submit_button)

    # find all necessary form elements to submit today's grade
    score_inputs = chrome.find_elements(
        '//input[@aria-valuemin="0" and @aria-valuemax="5"]')
    score_strengths = chrome.find_elements(
        '//div[@class="attitude-item positive"]')
    submit_button = chrome.find_element(
        '//button/span[contains(text(), "Submit")]/..')

    # loop through score inputs and give myself the highest score
    for score_input in score_inputs:
        chrome.type_into_element(score_input, '5')

    # loop through strenghts/weaknesses and give myself all strengths
    for score_strength in score_strengths:
        chrome.click_element(score_strength)

    # submit grade for today
    chrome.click_element(submit_button)
