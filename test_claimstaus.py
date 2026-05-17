import pytest
import pandas as pd
from playwright.sync_api import Page, expect, sync_playwright
from datetime import datetime
import shutil
from pathlib import Path
import time
import os
import logging




def test_login(page:Page):
    logging.basicConfig(level=logging.INFO)
    logging.info("Process started...")
    page.goto("https://login.zirmed.com/UI/Login")
    page.locator("id=loginName").fill("kd11")
    page.locator("id=password").fill("Optumisbest#432156")
    page.locator("id=loginButton").click()
    df=pd.read_excel("dev1.xlsx", sheet_name="ZirmedStatus")
    result=[]
    previous_charge = None
    for index, row in df.iterrows():
        logging.info("processing by bot 1")
        logging.info(f"Processing transaction item no: {index}")        
        account=row["Invoice"]
        logging.info(f"processing invoice no: {account} started...")
        user=row["User"]
        dos=pd.to_datetime(row["DOS"]).strftime("%m/%d/%Y")
        billed_amount=row["BilledAmount"]
        cpr_status=row["CPRStatus"]
        location=row["Location"]
        payorselection=row["PayorSelection"]
        site= page.locator(".header-account-search-text").input_value()
        if site==payorselection:
            pass
        else:
            page.get_by_title("Account Search").click()
            page.locator(".header-account-search-input").fill(payorselection)
            page.locator("#accountSearchChildButton").click()
            page.get_by_role("link", name=payorselection).click()
            element=page.get_by_text("Claims Processing")
            element.wait_for(state="visible")
            box=element.bounding_box()
            if box:
                page.mouse.move(box["x"] + box["width"]/2,box["y"] + box["height"]/2,steps=40)
                element.click()
            page.get_by_role("link", name="Claims", exact=True).click()
        page.locator("#SearchCriteria_Status").select_option(label="All")
        page.locator("#SearchCriteria_PatNumber").fill(str(account))
        page.locator("#ClaimListingSearchButtonTop").click()
        page.wait_for_timeout(1500)
        try:
            current_charges=page.locator("td.gridViewCell.chargesCell").first.text_content(timeout=2000).strip()
        except:
            current_charges=""
        
        if previous_charge is not None and current_charges == previous_charge:            
            page.wait_for_timeout(4000)
            try:
                current_charges=page.locator("td.gridViewCell.chargesCell").first.text_content(timeout=500).strip()
            except:
                current_charges=""        
        try:
            previous_charge=current_charges            
            status=page.locator("td.gridViewCell.descriptionCell").text_content(timeout=500).strip()
        except:
            status=" "
        logging.info(f"processing invoice no: {account} ended")
        result.append({"Invoice":account,"User":user,"DOS":dos,"Billed Amount":billed_amount,"CPRStatus":cpr_status,"Location":location,
                       "PayorSelection":payorselection,"ZermedStatus":status})
    result_df=pd.DataFrame(result)
    result_df.to_excel("output1.xlsx", index=False)


def test_login1(page:Page):
    logging.basicConfig(level=logging.INFO)
    logging.info("Process started...")
    page.goto("https://login.zirmed.com/UI/Login")
    page.locator("id=loginName").fill("kd11")
    page.locator("id=password").fill("Optumisbest#432156")
    page.locator("id=loginButton").click()
    df=pd.read_excel("dev2.xlsx", sheet_name="ZirmedStatus")
    result=[]
    previous_charge = None
    for index, row in df.iterrows():
        logging.info("processing by bot 2")
        logging.info(f"Processing transaction item no:{index}") 
        account=row["Invoice"]
        logging.info(f"processing invoice no: {account} started...")
        user=row["User"]
        dos=pd.to_datetime(row["DOS"]).strftime("%m/%d/%Y")
        billed_amount=row["BilledAmount"]
        cpr_status=row["CPRStatus"]
        location=row["Location"]
        payorselection=row["PayorSelection"]
        site= page.locator(".header-account-search-text").input_value()
        if site==payorselection:
            pass
        else:
            page.get_by_title("Account Search").click()
            page.locator(".header-account-search-input").fill(payorselection)
            page.locator("#accountSearchChildButton").click()
            page.get_by_role("link", name=payorselection).click()
            element=page.get_by_text("Claims Processing")
            element.wait_for(state="visible")
            box=element.bounding_box()
            if box:
                page.mouse.move(box["x"] + box["width"]/2,box["y"] + box["height"]/2,steps=40)
                element.click()
            page.get_by_role("link", name="Claims", exact=True).click()
        page.locator("#SearchCriteria_Status").select_option(label="All")
        page.locator("#SearchCriteria_PatNumber").fill(str(account))
        page.locator("#ClaimListingSearchButtonTop").click()
        page.wait_for_timeout(1500)
        try:
            current_charges=page.locator("td.gridViewCell.chargesCell").first.text_content(timeout=2000).strip()
        except:
            current_charges=""
        
        if previous_charge is not None and current_charges == previous_charge:            
            page.wait_for_timeout(4000)
            try:
                current_charges=page.locator("td.gridViewCell.chargesCell").first.text_content(timeout=500).strip()
            except:
                current_charges=""
        try:
            previous_charge=current_charges            
            status=page.locator("td.gridViewCell.descriptionCell").text_content(timeout=500).strip()
        except:
            status=" "
        logging.info(f"processing invoice no: {account} ended.")
        result.append({"Invoice":account,"User":user,"DOS":dos,"Billed Amount":billed_amount,"CPRStatus":cpr_status,"Location":location,
                       "PayorSelection":payorselection,"ZermedStatus":status})
    result_df=pd.DataFrame(result)
    result_df.to_excel("output2.xlsx", index=False)


    
    
