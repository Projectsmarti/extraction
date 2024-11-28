import pandas as pd
import google.generativeai as genai
import os
import re
from dotenv import load_dotenv
import logging


class AdvancedContentAnalyzer:
    def __init__(self):
        # Setup logging
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
        self.logger = logging.getLogger(__name__)

        # Setup Google AI
        load_dotenv()
        api_key = os.getenv('GOOGLE_API_KEY')
        if not api_key:
            raise ValueError("API key not found in .env file")

        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel('gemini-pro')

    def extract_contact_info(self, content: str) -> dict:
        """Extract contact information from the content"""
        phone_pattern = r'\b(?:\+\d{1,3}[-.\s]?)?(?:\(\d{3}\)|\d{3})[-.\s]?\d{3}[-.\s]?\d{4}\b'
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

        phones = re.findall(phone_pattern, content)
        emails = re.findall(email_pattern, content)

        return {
            'phones': list(set(phones)),
            'emails': list(set(emails))
        }

    def analyze_content(self, content: str) -> dict:
        """Analyze content using a comprehensive prompt"""
        prompt = f"""
        Thoroughly analyze this car service center description and extract detailed information:

        Description: {content}

        Categorize services into:
        1. Primary Services: Core, main services the center specializes in
        2. Secondary Services: Additional, supporting services
        3. Additional Services: Supplementary or optional services

        Provide the output in this structured format:
        Primary Services: [List primary services]
        Secondary Services: [List secondary services]
        Additional Services: [List additional services]
        """

        try:
            response = self.model.generate_content(prompt)
            return self._parse_detailed_response(response.text, content)
        except Exception as e:
            self.logger.error(f"Error in content analysis: {e}")
            return self._create_empty_result()

    def _parse_detailed_response(self, response: str, original_content: str) -> dict:
        """Parse the AI response into structured data with contact info"""
        result = self._create_empty_result()

        # Parse services
        service_sections = {
            'primary_services': 'Primary Services:',
            'secondary_services': 'Secondary Services:',
            'additional_services': 'Additional Services:'
        }

        for key, section_header in service_sections.items():
            section_match = re.search(
                f'{section_header}(.*?)(?=Primary Services:|Secondary Services:|Additional Services:|$)',
                response, re.DOTALL | re.IGNORECASE)
            if section_match:
                result[key] = section_match.group(1).strip().strip('[]').replace('\n', ', ')

        # Extract contact information
        contact_info = self.extract_contact_info(original_content)
        result['phones'] = ', '.join(contact_info['phones'])
        result['emails'] = ', '.join(contact_info['emails'])

        return result

    def _create_empty_result(self) -> dict:
        """Create empty result structure"""
        return {
            'primary_services': '',
            'secondary_services': '',
            'additional_services': '',
            'phones': '',
            'emails': ''
        }

    def process_file(self, input_file: str) -> pd.DataFrame:
        """Process the input file and analyze All_Content"""
        try:
            # Read the file
            df = pd.read_excel(input_file)
            self.logger.info(f"Loaded file with {len(df)} rows")

            # Initialize lists for new columns
            analyzed_data = []

            # Process each row
            for idx, row in df.iterrows():
                self.logger.info(f"Processing row {idx + 1}/{len(df)}")

                all_content = str(row.get('All_Content', ''))
                if all_content.strip():
                    analysis = self.analyze_content(all_content)
                else:
                    analysis = self._create_empty_result()

                # Combine original data with analysis
                processed_row = {
                    # Original columns to retain
                    'Name': row.get('Name', ''),
                    'Address': row.get('Address', ''),
                    'Rating_text': row.get('Rating_text', ''),
                    'Opening_Time': row.get('Opening_Time', ''),
                    'Phone_Number': row.get('Phone_Number', ''),
                    'GMAP': row.get('GMAP', ''),
                    'Website': row.get('Website', ''),

                    # Analyzed services
                    'Primary_Services': analysis['primary_services'],
                    'Secondary_Services': analysis['secondary_services'],
                    'Additional_Services': analysis['additional_services'],

                    # Extracted contact information
                    'Extracted_Phones': analysis['phones'],
                    'Extracted_Emails': analysis['emails']
                }
                analyzed_data.append(processed_row)

            # Create result DataFrame
            result_df = pd.DataFrame(analyzed_data)
            return result_df

        except Exception as e:
            self.logger.error(f"Error processing file: {e}")
            raise


def main():
    try:
        # Initialize analyzer
        analyzer = AdvancedContentAnalyzer()

        # Process file
        input_file = r"C:\Users\DELL\Downloads\scraped_2511.xlsx"  # Your Excel file path
        result_df = analyzer.process_file(input_file)

        # Save results
        output_file = "advanced_analyzed_services.xlsx"
        result_df.to_excel(output_file, index=False)
        print(f"Analysis completed. Results saved to {output_file}")

        # Display sample results
        print("\nSample of processed data:")
        print(result_df[['Name', 'Primary_Services', 'Secondary_Services', 'Additional_Services', 'Extracted_Phones',
                         'Extracted_Emails']].head())

    except Exception as e:
        print(f"Error in main execution: {e}")


if __name__ == "__main__":
    main()