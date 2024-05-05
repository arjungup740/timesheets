# Use the official AWS Lambda Python runtime as base image
FROM public.ecr.aws/lambda/python:3.8

# Copy requirements.txt into the container
COPY requirements.txt ./

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# copy excel template
COPY Arjun-BOI-Timesheet-04_20_24.xlsx ./

# Copy the lambda function code into the container
COPY lambda_function.py ./

# Set the CMD to your handler function
CMD ["lambda_function.lambda_handler"]
