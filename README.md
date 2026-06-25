# Alfa Laval HMI Image Reader

This Streamlit app extracts process values, valve status, pump status, and digital indicator status from Alfa Laval HMI screenshots.

## Features

- Upload one or multiple HMI screenshots
- Extract numerical process values
- Extract valve open/closed status
- Extract pump running/stopped status
- Record operator name and lot number
- Automatically reuse operator and lot number for 10 hours
- Save local Excel backup
- Upload data to SharePoint through Power Automate

## Required Streamlit Secrets

The app requires:

OPENAI_API_KEY  
POWER_AUTOMATE_URL