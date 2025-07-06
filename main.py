import streamlit as st
import tempfile
import os
from markitdown import MarkItDown
from pathlib import Path
import openai
from dotenv import load_dotenv

load_dotenv()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
if not OPENAI_API_KEY:
    st.error("❌ Missing OPENAI_API_KEY in environment or .env file")
    st.stop()

# Initialize OpenAI client
openai.api_key = OPENAI_API_KEY

def convert_docx_to_markdown(uploaded_file):
    """
    Convert uploaded DOCX file to markdown using markitdown library
    """
    try:
        # Create a temporary file to save the uploaded content
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        # Initialize MarkItDown converter
        md = MarkItDown()
        
        # Convert the file to markdown
        result = md.convert(tmp_file_path)
        
        # Clean up temporary file
        os.unlink(tmp_file_path)
        
        return result.text_content
    
    except Exception as e:
        return f"Error converting file: {str(e)}"

def convert_to_csv_with_openai(markdown_content):
    """
    Convert markdown content to CSV using OpenAI API
    """
    try:
        system_prompt = """
        Bạn là cán bộ chuyển đổi số về sáp nhập đơn vị quản lý hành chính ở Việt Nam.
        Công việc của bạn là đọc NGHỊ QUYẾT Về việc sắp xếp các đơn vị hành chính cấp xã của các tỉnh thành 
        rồi tổng hợp thành file CSV gồm các cột: Tỉnh, Hành chính cũ, Hành chính mới.

        Quy tắc:
        - Hành chính cũ: Phường/Xã/Thị xã/Thị trấn cũ trước khi chuyển đổi
        - Hành chính mới: Phường/Xã/Thị xã/Thị trấn mới sau khi chuyển đổi
        - Mỗi đơn vị hành chính cũ tạo thành 1 dòng riêng trong CSV
        - Chỉ trích xuất thông tin có trong nghị quyết, không suy diễn

        Ví dụ: "Sắp xếp toàn bộ diện tích tự nhiên, quy mô dân số của thị trấn Kiên Lương, xã Bình An (huyện Kiên Lương) và xã Bình Trị thành xã mới có tên gọi là xã Kiên Lương"
        
        Kết quả CSV:
        Tỉnh,Hành chính cũ,Hành chính mới
        An Giang,thị trấn Kiên Lương,xã Kiên Lương
        An Giang,xã Bình An,xã Kiên Lương
        An Giang,xã Bình Trị,xã Kiên Lương

        Hãy trả về CHỈ nội dung CSV, không có text giải thích gì khác.
        """

        user_prompt = f"""
        Đây là nội dung nghị quyết về việc sắp xếp các đơn vị hành chính cấp xã của các tỉnh thành.
        Hãy tổng hợp thành file CSV theo format đã yêu cầu:

        Nội dung nghị quyết:
        {markdown_content}
        """

        # Make API call to OpenAI
        response = openai.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.1,
            max_tokens=4000
        )
        
        csv_result = response.choices[0].message.content.strip()
        return csv_result

    except Exception as e:
        return f"Error calling OpenAI API: {str(e)}"

def quality_check_csv(markdown_content, csv_result):
    """
    Quality check the CSV result against the original markdown
    """
    try:
        system_prompt = """
        Bạn là cán bộ thanh tra về sáp nhập đơn vị quản lý hành chính ở Việt Nam.
        Công việc của bạn là kiểm tra file CSV đối chiếu với nội dung nghị quyết và đánh giá từng dòng.

        Quy tắc đánh giá:
        - Kiểm tra từng dòng CSV xem có chính xác với nội dung nghị quyết không
        - Thêm cột "Kết quả" vào cuối mỗi dòng
        - Ghi "Đúng" nếu thông tin chính xác, "Sai" nếu không chính xác
        - Chỉ dựa vào thông tin trong nghị quyết, không suy diễn

        Trả về CSV đã được bổ sung cột "Kết quả", không có text giải thích gì khác.
        """

        user_prompt = f"""
        Đây là nội dung nghị quyết gốc:
        {markdown_content}

        Và đây là file CSV cần kiểm tra:
        {csv_result}

        Hãy kiểm tra và thêm cột "Kết quả" vào CSV:
        """

        # Make API call to OpenAI
        response = openai.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.1,
            max_tokens=4000
        )
        
        qc_result = response.choices[0].message.content.strip()
        return qc_result

    except Exception as e:
        return f"Error in quality check: {str(e)}"

def main():
    st.set_page_config(
        page_title="DOCX to CSV Converter",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("📄 DOCX to CSV Converter")
    st.markdown("Upload a DOCX file and convert it to CSV format using OpenAI")
    
    # File uploader
    uploaded_file = st.file_uploader(
        "Choose a DOCX file",
        type=['docx'],
        help="Upload a Microsoft Word document (.docx) to convert to CSV"
    )
    
    if uploaded_file is not None:
        # Display file details
        st.success(f"File uploaded: {uploaded_file.name}")
        st.info(f"File size: {uploaded_file.size:,} bytes")
        
        # Convert button
        if st.button("Convert to CSV", type="primary"):
            # Step 1: Convert DOCX to Markdown
            with st.spinner("Converting DOCX to Markdown..."):
                markdown_content = convert_docx_to_markdown(uploaded_file)
            
            if markdown_content.startswith("Error"):
                st.error(markdown_content)
                return
            
            st.success("✅ Markdown conversion completed!")
            
            # Step 2: Convert Markdown to CSV using OpenAI
            with st.spinner("Converting to CSV using OpenAI..."):
                csv_result = convert_to_csv_with_openai(markdown_content)
            
            if csv_result.startswith("Error"):
                st.error(csv_result)
                return
            
            st.success("✅ CSV conversion completed!")
            
            # Step 3: Quality check (optional)
            with st.spinner("Performing quality check..."):
                qc_result = quality_check_csv(markdown_content, csv_result)
            
            if qc_result.startswith("Error"):
                st.warning(f"Quality check failed: {qc_result}")
                final_csv = csv_result
            else:
                st.success("✅ Quality check completed!")
                final_csv = qc_result
            
            # Display results
            col1, col2, col3 = st.columns([1, 1, 1])
            
            with col1:
                st.subheader("📝 Original Markdown")
                st.text_area(
                    "Raw Markdown",
                    value=markdown_content[:2000] + "..." if len(markdown_content) > 2000 else markdown_content,
                    height=400,
                    help="Original markdown content (truncated for display)"
                )
            
            with col2:
                st.subheader("📊 Generated CSV")
                st.text_area(
                    "CSV Result",
                    value=csv_result,
                    height=400,
                    help="CSV generated by OpenAI"
                )
                
                # Download button for original CSV
                st.download_button(
                    label="💾 Download Original CSV",
                    data=csv_result,
                    file_name=f"{Path(uploaded_file.name).stem}_original.csv",
                    mime="text/csv",
                    help="Download the original CSV file"
                )
            
            with col3:
                st.subheader("✅ Quality Checked CSV")
                st.text_area(
                    "QC Result",
                    value=final_csv,
                    height=400,
                    help="CSV with quality check results"
                )
                
                # Download button for QC CSV
                st.download_button(
                    label="💾 Download Final CSV",
                    data=final_csv,
                    file_name=f"{Path(uploaded_file.name).stem}_final.csv",
                    mime="text/csv",
                    help="Download the quality-checked CSV file"
                )
            
            # Show CSV preview as table
            st.subheader("📋 CSV Preview")
            try:
                # Parse CSV for display
                lines = final_csv.strip().split('\n')
                if len(lines) > 0:
                    headers = lines[0].split(',')
                    rows = [line.split(',') for line in lines[1:]]
                    
                    # Create DataFrame-like display
                    import pandas as pd
                    df = pd.DataFrame(rows, columns=headers)
                    st.dataframe(df, use_container_width=True)
                else:
                    st.warning("No data to display")
            except Exception as e:
                st.error(f"Error displaying CSV preview: {str(e)}")
    
    else:
        st.info("👆 Please upload a DOCX file to get started")
    
    # Footer
    st.markdown("---")
    st.markdown("**Note:** From Linh Nguyen - Credit PM - ShopeePay")

if __name__ == "__main__":
    main()