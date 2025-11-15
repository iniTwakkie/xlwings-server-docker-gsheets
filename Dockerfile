FROM xlwings/xlwings-server:0.7.0
RUN pip install --no-cache-dir requests
COPY ./custom_functions /project/app/custom_functions
COPY ./custom_scripts /project/app/custom_scripts