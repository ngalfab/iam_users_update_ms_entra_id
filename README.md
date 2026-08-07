# iam_users_update_ms_entra_id

deploying ci/cd pipeline to create user account in MS Entra ID

## Entra ID Users Management and Automation pipeline

This repository provides an automated CI/CD framework to manage and update Microsoft Entra ID uwers in bulk. It includes dataset, Graph API integration using python, and github action workflow to automate directory.

## Directory Structure

.
├── .github/
│ └── workflows/
│ └── update-entra-users.yml # GitHub Actions automation pipeline
├── sc300_practice_users_dallas-v2.csv # Sample dataset with Dallas locality
├── update_users.py # Python script using MSAL and Graph API
└── README.md # Documentation
