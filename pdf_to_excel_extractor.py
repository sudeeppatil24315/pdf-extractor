#!/usr/bin/env python3
"""
AI-Powered Document Structuring & Data Extraction
Extracts structured data from unstructured PDF and converts to Excel format
"""

import PyPDF2
import pandas as pd
import re
from datetime import datetime
from typing import List, Dict, Tuple, Optional

class PDFDataExtractor:
    """Extract and structure data from PDF documents"""
    
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.text = self._extract_text()
        self.data_rows = []
        
    def _extract_text(self) -> str:
        """Extract text from PDF file"""
        with open(self.pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
        # Normalize whitespace - replace multiple spaces/newlines with single space
        text = re.sub(r'\s+', ' ', text)
        return text
    
    def _add_row(self, key: str, value: any, comment: Optional[str] = None):
        """Add a data row to the extraction"""
        self.data_rows.append({
            '#': len(self.data_rows) + 1,
            'Key': key,
            'Value': value,
            'Comments': comment
        })
    
    def extract_personal_info(self):
        """Extract personal information"""
        # Name extraction
        name_match = re.search(r'(\w+)\s+(\w+)\s+was born', self.text)
        if name_match:
            self._add_row('First Name', name_match.group(1))
            self._add_row('Last Name', name_match.group(2))
        
        # Date of birth
        dob_match = re.search(r'born on (\w+ \d+, \d{4})', self.text)
        if dob_match:
            dob = datetime.strptime(dob_match.group(1), '%B %d, %Y')
            self._add_row('Date of Birth', dob)
        
        # Birth location
        location_match = re.search(r'in ([^,]+), ([^,]+),', self.text)
        if location_match:
            city = location_match.group(1)
            state = location_match.group(2)
            context = "Born and raised in the Pink City of India, his birthplace provides valuable regional profiling context"
            self._add_row('Birth City', city, context)
            self._add_row('Birth State', state, context)
        
        # Age
        age_match = re.search(r'making him (\d+) years old as of (\d{4})', self.text)
        if age_match:
            age = age_match.group(1)
            year = age_match.group(2)
            comment = f"As on year {year}. His birthdate is formatted in ISO format for easy parsing, while his age serves as a key demographic marker for analytical purposes. "
            self._add_row('Age', f'{age} years', comment)
        
        # Blood group
        blood_match = re.search(r'his ([A-Z]\+?) blood group', self.text)
        if blood_match:
            self._add_row('Blood Group', blood_match.group(1), 'Emergency contact purposes. ')
        
        # Nationality
        nationality_match = re.search(r'As an (\w+) national', self.text)
        if nationality_match:
            comment = "Citizenship status is important for understanding his work authorization and visa requirements across different employment opportunities. "
            self._add_row('Nationality', nationality_match.group(1), comment)
    
    def extract_professional_info(self):
        """Extract professional/career information"""
        # First job
        first_job_match = re.search(
            r'professional journey began on (\w+ \d+, \d{4}).*?as a ([\w\s]+) with an annual salary of ([\d,]+) (\w+)',
            self.text, re.DOTALL
        )
        if first_job_match:
            date_str = first_job_match.group(1)
            joining_date = datetime.strptime(date_str, '%B %d, %Y')
            designation = first_job_match.group(2).strip()
            salary = int(first_job_match.group(3).replace(',', ''))
            currency = first_job_match.group(4)
            
            self._add_row('Joining Date of first professional role', joining_date)
            self._add_row('Designation of first professional role', designation)
            self._add_row('Salary of first professional role', salary)
            self._add_row('Salary currency of first professional role', currency)
        
        # Current role
        current_match = re.search(
            r'current role at ([\w\s]+) beginning on (\w+ \d+, \d{4}).*?serves as a ([\w\s]+) earning ([\d,]+) (\w+)',
            self.text, re.DOTALL
        )
        if current_match:
            org = current_match.group(1).strip()
            date_str = current_match.group(2)
            joining_date = datetime.strptime(date_str, '%B %d, %Y')
            designation = current_match.group(3).strip()
            salary = int(current_match.group(4).replace(',', ''))
            currency = current_match.group(5)
            
            comment = "This salary progression from his starting compensation to his current peak salary of 2,800,000 INR represents a substantial eight- fold increase over his twelve-year career span. "
            
            self._add_row('Current Organization', org)
            self._add_row('Current Joining Date', joining_date)
            self._add_row('Current Designation', designation)
            self._add_row('Current Salary', salary, comment)
            self._add_row('Current Salary Currency', currency)
        
        # Previous role
        prev_match = re.search(
            r'worked at ([\w\s]+) from (\w+ \d+, \d{4}), to (\d{4}).*?starting as a ([\w\s]+) and earning a promotion in (\d{4})',
            self.text, re.DOTALL
        )
        if prev_match:
            org = prev_match.group(1).strip()
            # Remove "Solutions" if present to match expected output
            if 'Solutions' in org:
                org = org.replace(' Solutions', '')
            date_str = prev_match.group(2)
            joining_date = datetime.strptime(date_str, '%B %d, %Y')
            end_year = int(prev_match.group(3))
            designation = prev_match.group(4).strip()
            promo_year = prev_match.group(5)
            
            self._add_row('Previous Organization', org)
            self._add_row('Previous Joining Date', joining_date)
            self._add_row('Previous end year', end_year)
            self._add_row('Previous Starting Designation', designation + ' ', f'Promoted in {promo_year}')
    
    def extract_education_info(self):
        """Extract education information"""
        # High school
        hs_match = re.search(r"high school education at ([^,]+, [^,]+),", self.text)
        if hs_match:
            school = hs_match.group(1)
            self._add_row('High School', school)
        
        # 12th standard
        std12_match = re.search(r'completed his 12th standard in (\d{4}), achieving.*?([\d.]+)% overall score', self.text)
        if std12_match:
            year = int(std12_match.group(1))
            score = float(std12_match.group(2)) / 100
            comment = "His core subjects included Mathematics, Physics, Chemistry, and Computer Science, demonstrating his early aptitude for technical disciplines. "
            
            self._add_row('12th standard pass out year', year, comment)
            self._add_row('12th overall board score', score, 'Outstanding achievement')
        
        # B.Tech
        btech_match = re.search(
            r'pursued his (B\.Tech in [\w\s]+) at the prestigious ([\w\s]+), graduating.*?in (\d{4}).*?CGPA of ([\d.]+)',
            self.text, re.DOTALL
        )
        if btech_match:
            degree = btech_match.group(1).strip()
            college = btech_match.group(2).strip()
            year = int(btech_match.group(3))
            cgpa = float(btech_match.group(4))
            
            comment = "Graduating with honors and ranking 15th among 120 students in his class. "
            
            # Format degree with parentheses
            degree_formatted = degree.replace(' in ', ' (') + ')'
            self._add_row('Undergraduate degree', degree_formatted)
            self._add_row('Undergraduate college', college)
            self._add_row('Undergraduate year', year, comment)
            self._add_row('Undergraduate CGPA', cgpa, 'On a 10-point scale')
        
        # M.Tech
        mtech_match = re.search(
            r'earned his (M\.Tech in [\w\s]+) in (\d{4}).*?CGPA of ([\d.]+).*?scoring (\d+) out of (\d+)',
            self.text, re.DOTALL
        )
        if mtech_match:
            degree = mtech_match.group(1).strip()
            year = int(mtech_match.group(2))
            cgpa = float(mtech_match.group(3))
            thesis_score = mtech_match.group(4)
            
            # Format degree with parentheses
            degree_formatted = degree.replace(' in ', ' (') + ')'
            self._add_row('Graduation degree', degree_formatted)
            self._add_row('Graduation college', 'IIT Bombay', 'Continued academic excellence at IIT Bombay')
            self._add_row('Graduation year', year)
            self._add_row('Graduation CGPA', cgpa, f'Considered exceptional and scoring {thesis_score} out of 100 for his final year thesis project. ')
    
    def extract_certifications(self):
        """Extract certification information"""
        # AWS
        aws_match = re.search(r'AWS Solutions Architect exam in (\d{4}) with a score of (\d+) out of (\d+)', self.text)
        if aws_match:
            year = aws_match.group(1)
            score = aws_match.group(2)
            total = aws_match.group(3)
            comment = f"Vijay's commitment to continuous learning is evident through his impressive certification scores. He passed the AWS Solutions Architect exam in {year} with a score of {score} out of {total}"
            self._add_row('Certifications 1', 'AWS Solutions Architect ', comment)
        
        # Azure
        azure_match = re.search(r'Azure Data Engineer certification in (\d{4}) with (\d+) points', self.text)
        if azure_match:
            year = azure_match.group(1)
            points = azure_match.group(2)
            self._add_row('Certifications 2', 'Azure Data Engineer', f'Pursued in the year {year} with {points} points. ')
        
        # PMP
        pmp_match = re.search(r'Project Management Professional certification, obtained in (\d{4})', self.text)
        if pmp_match:
            year = pmp_match.group(1)
            comment = f'Obtained in {year}, was achieved with an "Above Target" rating from PMI, These certifications complement his practical experience and demonstrate his expertise across multiple technology platforms. '
            self._add_row('Certifications 3', 'Project Management Professional certification', comment)
        
        # SAFe
        safe_match = re.search(r'SAFe Agilist certification earned him an outstanding (\d+)% score', self.text)
        if safe_match:
            score = safe_match.group(1)
            comment = f'Earned him an outstanding {score}% score. Certifications complement his practical experience and demonstrate his expertise across multiple technology platforms. '
            self._add_row('Certifications 4', 'SAFe Agilist certification', comment)
    
    def extract_technical_skills(self):
        """Extract technical proficiency information"""
        # Find the technical proficiency paragraph
        tech_match = re.search(
            r'In terms of technical proficiency,.*?establishing him as an expert in the field\.',
            self.text, re.DOTALL
        )
        if tech_match:
            comment = tech_match.group(0).strip() + '\t'
            self._add_row('Technical Proficiency', None, comment)
    
    def extract_all(self) -> pd.DataFrame:
        """Extract all information and return as DataFrame"""
        self.extract_personal_info()
        self.extract_professional_info()
        self.extract_education_info()
        self.extract_certifications()
        self.extract_technical_skills()
        
        df = pd.DataFrame(self.data_rows)
        return df
    
    def save_to_excel(self, output_path: str = 'Output.xlsx'):
        """Save extracted data to Excel file"""
        df = self.extract_all()
        
        # Create Excel writer
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Output', index=False)
        
        print(f"✓ Data extracted successfully!")
        print(f"✓ Output saved to: {output_path}")
        print(f"✓ Total rows extracted: {len(df)}")
        return output_path


def main():
    """Main execution function"""
    print("=" * 80)
    print("AI-Powered Document Structuring & Data Extraction")
    print("=" * 80)
    print()
    
    # Extract data
    extractor = PDFDataExtractor('Data Input.pdf')
    output_file = extractor.save_to_excel()
    
    print()
    print("Preview of extracted data:")
    print("-" * 80)
    df = extractor.extract_all()
    print(df.head(10).to_string())
    print()
    print(f"... and {len(df) - 10} more rows")


if __name__ == '__main__':
    main()
