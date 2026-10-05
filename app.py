import streamlit as st
from pypdf import PdfReader
from docx import Document
from transformers import AutoModelForSeq2SeqLM, AutoTokenizer


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="SummarAI",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CUSTOM STYLING
# ============================================================

st.markdown(
    """
    <style>

    /* ========================================================
       MAIN APP
       ======================================================== */

    .stApp {
        background: linear-gradient(
            135deg,
            #faf7ff 0%,
            #f5effb 50%,
            #eee7f7 100%
        );
    }

    .main .block-container {
        max-width: 1250px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }


    /* ========================================================
       MAIN HEADINGS
       ======================================================== */

    [data-testid="stAppViewContainer"] h1 {
        color: #32194d !important;
        -webkit-text-fill-color: #32194d !important;
        font-weight: 800 !important;
    }

    [data-testid="stAppViewContainer"] h2 {
        color: #382052 !important;
        -webkit-text-fill-color: #382052 !important;
        font-weight: 750 !important;
    }

    [data-testid="stAppViewContainer"] h3 {
        color: #382052 !important;
        -webkit-text-fill-color: #382052 !important;
        font-weight: 700 !important;
    }


    /* ========================================================
       MAIN TEXT
       ======================================================== */

    [data-testid="stAppViewContainer"] .stMarkdown p {
        color: #352348 !important;
        -webkit-text-fill-color: #352348 !important;
        font-weight: 500 !important;
    }

    [data-testid="stAppViewContainer"] .stMarkdown strong {
        color: #2f1947 !important;
        -webkit-text-fill-color: #2f1947 !important;
        font-weight: 700 !important;
    }


    /* ========================================================
       SIDEBAR
       ======================================================== */

    [data-testid="stSidebar"] {
        background: linear-gradient(
            180deg,
            #2b1949 0%,
            #39235f 55%,
            #45296e 100%
        );
    }

    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] label {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }


    /* ========================================================
       SIDEBAR NORMAL TEXT
       ======================================================== */

    [data-testid="stSidebar"] .stMarkdown p {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }


    /* ========================================================
       SIDEBAR SELECT BOX
       ======================================================== */

    [data-testid="stSidebar"] [data-baseweb="select"] > div {
        background-color: #ffffff !important;
        border-radius: 10px !important;
    }


    /* ========================================================
       SIDEBAR RADIO BUTTONS
       ======================================================== */

    [data-testid="stSidebar"] [role="radiogroup"] label {
        color: #ffffff !important;
    }

    [data-testid="stSidebar"] [role="radiogroup"] label p {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        font-weight: 600 !important;
    }


    /* ========================================================
       SIDEBAR INFO / SUCCESS BOXES
       ======================================================== */

    [data-testid="stSidebar"] [data-testid="stAlert"] {
        border-radius: 12px !important;
    }

    [data-testid="stSidebar"] [data-testid="stAlert"] p {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        font-weight: 600 !important;
    }


    /* ========================================================
       IMPORTANT:
       SIDEBAR FOOTER TEXT
       ======================================================== */

    [data-testid="stSidebar"] .stMarkdown p.sidebar-footer-text {
        color: #f0e6ff !important;
        -webkit-text-fill-color: #f0e6ff !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        opacity: 1 !important;
        margin-top: 0 !important;
        margin-bottom: 8px !important;
        visibility: visible !important;
    }


    /* ========================================================
       MAIN SELECT BOX
       ======================================================== */

    [data-testid="stAppViewContainer"]
    [data-testid="stSelectbox"]
    [data-baseweb="select"] > div {
        background-color: #ffffff !important;
        border: 2px solid #9a78b8 !important;
        border-radius: 11px !important;
        min-height: 45px !important;
    }


    /* ========================================================
       TEXT AREAS
       ======================================================== */

    textarea {
        color: #2d193d !important;
        -webkit-text-fill-color: #2d193d !important;
        background-color: #ffffff !important;
        font-size: 15px !important;
        line-height: 1.6 !important;
        border-radius: 12px !important;
        opacity: 1 !important;
    }

    textarea:disabled {
        color: #2d193d !important;
        -webkit-text-fill-color: #2d193d !important;
        background-color: #ffffff !important;
        opacity: 1 !important;
    }


    /* ========================================================
       FILE UPLOADER
       ======================================================== */

    [data-testid="stFileUploader"] {
        background-color: #ffffff !important;
        border-radius: 14px !important;
        padding: 10px !important;
        border: 1px solid #ded2e9 !important;
    }


    /* ========================================================
       METRICS
       ======================================================== */

    [data-testid="stMetric"] {
        background-color: #ffffff !important;
        border-radius: 14px !important;
        padding: 15px !important;
        border: 1px solid #e2d8ed !important;
        box-shadow: 0 4px 14px rgba(54, 32, 80, 0.07);
    }

    [data-testid="stMetricLabel"] {
        color: #735a8a !important;
    }

    [data-testid="stMetricValue"] {
        color: #32194d !important;
        -webkit-text-fill-color: #32194d !important;
        font-weight: 800 !important;
    }


    /* ========================================================
       BUTTONS
       ======================================================== */

    .stButton > button {
        border-radius: 11px !important;
        font-weight: 700 !important;
        min-height: 45px !important;
    }

    .stButton > button[kind="primary"] {
        background: linear-gradient(
            90deg,
            #6d3bb3,
            #a054ca
        ) !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        border: none !important;
    }

    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(
            90deg,
            #5d2fa0,
            #8e45b8
        ) !important;
    }

    .stDownloadButton > button {
        background-color: #ffffff !important;
        color: #57327d !important;
        -webkit-text-fill-color: #57327d !important;
        border: 1px solid #9b7bbb !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
    }


    /* ========================================================
       ALERTS
       ======================================================== */

    [data-testid="stAlert"] {
        border-radius: 12px !important;
    }


    /* ========================================================
       DIVIDERS
       ======================================================== */

    hr {
        border-color: #ddd1e8 !important;
    }


    /* ========================================================
       CAPTIONS
       ======================================================== */

    [data-testid="stCaptionContainer"] p {
        color: #735d82 !important;
        -webkit-text-fill-color: #735d82 !important;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SESSION STATE
# ============================================================

if "original_text" not in st.session_state:
    st.session_state.original_text = ""

if "summary_text" not in st.session_state:
    st.session_state.summary_text = ""


# ============================================================
# AI MODELS
# ============================================================

MODELS = {
    "DistilBART — Fast": "sshleifer/distilbart-cnn-12-6",
    "BART Large — Quality": "facebook/bart-large-cnn",
}


# ============================================================
# SUMMARY LENGTH
# ============================================================

LENGTHS = {
    "Short": (30, 80),
    "Medium": (60, 150),
    "Long": (100, 250),
}

CHUNK_WORDS = 600


# ============================================================
# LOAD MODEL
# ============================================================

@st.cache_resource
def load_model(model_name):

    tokenizer = AutoTokenizer.from_pretrained(
        model_name
    )

    model = AutoModelForSeq2SeqLM.from_pretrained(
        model_name
    )

    return tokenizer, model


# ============================================================
# CHUNK TEXT
# ============================================================

def chunk_text(text, max_words=CHUNK_WORDS):

    words = text.split()

    chunks = []

    for i in range(
        0,
        len(words),
        max_words
    ):

        chunk = " ".join(
            words[i:i + max_words]
        )

        if chunk.strip():
            chunks.append(chunk)

    return chunks


# ============================================================
# SUMMARIZE ONE CHUNK
# ============================================================

def summarize_chunk(
    text,
    tokenizer,
    model,
    min_length,
    max_length
):

    inputs = tokenizer(
        text,
        return_tensors="pt",
        truncation=True,
        max_length=1024,
    )

    summary_ids = model.generate(
        inputs["input_ids"],
        attention_mask=inputs["attention_mask"],
        min_length=min_length,
        max_length=max_length,
        num_beams=2,
        length_penalty=2.0,
        no_repeat_ngram_size=3,
        early_stopping=True,
    )

    return tokenizer.decode(
        summary_ids[0],
        skip_special_tokens=True,
    )


# ============================================================
# SUMMARIZE DOCUMENT
# ============================================================

def summarize_document(
    text,
    tokenizer,
    model,
    min_length,
    max_length
):

    chunks = chunk_text(text)

    if not chunks:
        return ""

    progress = st.progress(0)

    status = st.empty()

    summaries = []

    for index, chunk in enumerate(chunks):

        status.write(
            f"🧠 Summarizing section "
            f"{index + 1} of {len(chunks)}..."
        )

        summary = summarize_chunk(
            chunk,
            tokenizer,
            model,
            min_length,
            max_length,
        )

        if summary.strip():
            summaries.append(summary)

        progress.progress(
            int(
                ((index + 1) / len(chunks))
                * 100
            )
        )

    status.write(
        "✨ Finalizing your summary..."
    )

    combined_summary = " ".join(
        summaries
    )

    if len(combined_summary.split()) > max_length * 2:

        combined_summary = summarize_chunk(
            combined_summary,
            tokenizer,
            model,
            min_length,
            max_length,
        )

    progress.empty()
    status.empty()

    return combined_summary


# ============================================================
# READ PDF
# ============================================================

def read_pdf(uploaded_file):

    reader = PdfReader(
        uploaded_file
    )

    pages = []

    for page in reader.pages:

        text = page.extract_text()

        if text:
            pages.append(text)

    return "\n".join(pages)


# ============================================================
# READ DOCX
# ============================================================

def read_docx(uploaded_file):

    document = Document(
        uploaded_file
    )

    paragraphs = []

    for paragraph in document.paragraphs:

        if paragraph.text.strip():
            paragraphs.append(
                paragraph.text
            )

    return "\n".join(paragraphs)


# ============================================================
# READ TXT
# ============================================================

def read_txt(uploaded_file):

    return uploaded_file.read().decode(
        "utf-8",
        errors="ignore",
    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.title("⚙️ Settings")

    st.markdown("### 🤖 AI Model")

    selected_model = st.selectbox(
        "Choose summarization model",
        list(MODELS.keys()),
        index=0,
    )

    st.markdown("### 📏 Summary Length")

    selected_length = st.radio(
        "Choose output length",
        list(LENGTHS.keys()),
        index=0,
    )

    st.divider()

    st.markdown("### 💡 Recommended")

    st.info(
        "🚀 DistilBART is recommended for "
        "faster summaries and everyday use."
    )

    st.markdown("### 🔒 Local Processing")

    st.success(
        "🛡️ Your uploaded document is processed "
        "locally by the application."
    )

    st.divider()

    # --------------------------------------------------------
    # SIDEBAR FOOTER
    # --------------------------------------------------------

    st.markdown(
        '<p class="sidebar-footer-text">✨ SummarAI</p>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<p class="sidebar-footer-text">'
        'Intelligent Text Summarization System'
        '</p>',
        unsafe_allow_html=True,
    )


# ============================================================
# MAIN HEADER
# ============================================================

st.title("✨ SummarAI")

st.subheader(
    "✨ Intelligent Text Summarization System"
)

st.markdown(
    "**Transform long documents into concise, readable "
    "summaries using AI-powered transformer models.**"
)


# ============================================================
# FEATURE CARDS
# ============================================================

feature1, feature2, feature3 = st.columns(3)

with feature1:

    st.info(
        "📄 **Multiple Formats**\n\n"
        "Summarize text, TXT, PDF, and DOCX files."
    )

with feature2:

    st.info(
        "🤖 **AI Powered**\n\n"
        "Use transformer-based summarization models."
    )

with feature3:

    st.info(
        "⚡ **Flexible Output**\n\n"
        "Choose short, medium, or long summaries."
    )


st.divider()


# ============================================================
# INPUT DOCUMENT
# ============================================================

st.subheader(
    "📥 Input Document"
)

st.markdown(
    "**Choose how you want to provide your document:**"
)


input_type = st.selectbox(
    "Input method",
    [
        "📋 Paste Text",
        "📂 Upload File",
    ],
    index=0,
    label_visibility="collapsed",
)


# ============================================================
# PASTE TEXT
# ============================================================

if input_type == "📋 Paste Text":

    st.markdown(
        "**✍️ Enter your text:**"
    )

    text_input = st.text_area(
        "Text to summarize",
        height=300,
        placeholder=(
            "Paste the article, notes, report, "
            "or other text you want to summarize here..."
        ),
        label_visibility="collapsed",
    )

    if text_input.strip():

        st.session_state.original_text = text_input


# ============================================================
# UPLOAD FILE
# ============================================================

else:

    st.markdown(
        "**📂 Upload your document:**"
    )

    uploaded_file = st.file_uploader(
        "Document file",
        type=[
            "txt",
            "pdf",
            "docx",
        ],
        label_visibility="collapsed",
    )

    if uploaded_file is not None:

        file_name = uploaded_file.name.lower()

        try:

            if file_name.endswith(".pdf"):

                extracted_text = read_pdf(
                    uploaded_file
                )

            elif file_name.endswith(".docx"):

                extracted_text = read_docx(
                    uploaded_file
                )

            elif file_name.endswith(".txt"):

                extracted_text = read_txt(
                    uploaded_file
                )

            else:

                extracted_text = ""

            if extracted_text.strip():

                st.session_state.original_text = (
                    extracted_text
                )

                st.success(
                    f"✅ Successfully loaded "
                    f"{uploaded_file.name}"
                )

            else:

                st.warning(
                    "⚠️ No readable text was found "
                    "in this file."
                )

        except Exception as error:

            st.error(
                f"❌ Could not read the uploaded file: "
                f"{error}"
            )


# ============================================================
# DOCUMENT INFORMATION
# ============================================================

if st.session_state.original_text.strip():

    word_count = len(
        st.session_state.original_text.split()
    )

    st.subheader(
        "📊 Document Information"
    )

    stat1, stat2, stat3 = st.columns(3)

    with stat1:

        st.metric(
            "📝 Words",
            f"{word_count:,}",
        )

    with stat2:

        st.metric(
            "🔤 Characters",
            f"{len(st.session_state.original_text):,}",
        )

    with stat3:

        reading_time = max(
            1,
            round(word_count / 200),
        )

        st.metric(
            "⏱️ Reading Time",
            f"{reading_time} min",
        )


# ============================================================
# ACTION BUTTONS
# ============================================================

st.divider()

button1, button2 = st.columns(
    [3, 1]
)

with button1:

    generate_button = st.button(
        "✨ Generate AI Summary",
        type="primary",
        use_container_width=True,
    )

with button2:

    clear_button = st.button(
        "🗑️ Clear",
        use_container_width=True,
    )


# ============================================================
# CLEAR
# ============================================================

if clear_button:

    st.session_state.original_text = ""
    st.session_state.summary_text = ""

    st.rerun()


# ============================================================
# GENERATE SUMMARY
# ============================================================

if generate_button:

    if not st.session_state.original_text.strip():

        st.warning(
            "⚠️ Please enter some text or upload "
            "a document first."
        )

    else:

        try:

            model_name = MODELS[
                selected_model
            ]

            min_length, max_length = LENGTHS[
                selected_length
            ]

            with st.spinner(
                "🤖 Loading AI model..."
            ):

                tokenizer, model = load_model(
                    model_name
                )

            summary = summarize_document(
                st.session_state.original_text,
                tokenizer,
                model,
                min_length,
                max_length,
            )

            if summary.strip():

                st.session_state.summary_text = summary

                st.success(
                    "🎉 Summary generated successfully!"
                )

            else:

                st.warning(
                    "⚠️ The model could not generate "
                    "a summary."
                )

        except Exception as error:

            st.error(
                f"❌ Something went wrong: {error}"
            )


# ============================================================
# SUMMARY RESULTS
# ============================================================

if st.session_state.summary_text.strip():

    st.divider()

    st.subheader(
        "✨ Summary Results"
    )

    original_words = len(
        st.session_state.original_text.split()
    )

    summary_words = len(
        st.session_state.summary_text.split()
    )

    if original_words > 0:

        reduction = round(
            (
                1
                - summary_words / original_words
            )
            * 100
        )

    else:

        reduction = 0


    # ========================================================
    # RESULT METRICS
    # ========================================================

    result1, result2, result3, result4 = st.columns(4)

    with result1:

        st.metric(
            "📄 Original",
            f"{original_words:,} words",
        )

    with result2:

        st.metric(
            "✨ Summary",
            f"{summary_words:,} words",
        )

    with result3:

        st.metric(
            "📉 Reduction",
            f"{reduction}%",
        )

    with result4:

        st.metric(
            "🤖 Model",
            selected_model.split(" — ")[0],
        )


    # ========================================================
    # ORIGINAL / SUMMARY
    # ========================================================

    left, right = st.columns(2)

    with left:

        st.markdown(
            "### 📄 Original Document"
        )

        st.text_area(
            "Original document",
            value=st.session_state.original_text,
            height=450,
            disabled=True,
            label_visibility="collapsed",
        )


    with right:

        st.markdown(
            "### ✨ AI Summary"
        )

        st.text_area(
            "Generated summary",
            value=st.session_state.summary_text,
            height=450,
            disabled=True,
            label_visibility="collapsed",
        )

        st.download_button(
            "⬇️ Download Summary",
            data=st.session_state.summary_text,
            file_name="summarai_summary.txt",
            mime="text/plain",
            use_container_width=True,
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "✨ SummarAI • Intelligent Text Summarization System • "
    "Powered by Streamlit & Hugging Face"
)