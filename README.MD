# Hot-Dog-Detection Using YOLOv8

This project is a solution for detecting hot dogs in images using YOLOv8 object detection mode. The model is trained to detect hot dogs in images by drawing boxes around them.

# Method

As mentioned, the project uses a pre-trained YOLOv8 (You Only Look Once) model to a dataset of hotdog images. YOLOv8 is the most modern and utilized method of detecting objects in images. 

# Steps

1. Data preparation:
- The images and labels of the dataset are loaded from a specified directories and prepared for training.

2. Training the Model:
- A YOLOv8 model (pretrained yolov8n.pt) is used on the hotdog dataset.
- The training process runs for 5 epochs with a batch size of 16 and image size of 640x640. See more info on this in the next steps section below.


3. Validation:
- The model is validated on the validation dataset to assess accuracy using object detection matrics (mAP)

4. Finale:
- After validation, a single image is shown to see the results in action.
- The model is saved as hotdog_detector.pt

# Required Libraries
- Pytorch
- OpenCV
- Matplotlib
- Ultralyrics
- Pathlib

# How to Run
1. Download file and dataset
2. Prepare dataset. Place the correct directories where your data is and update the data.yaml file
3. Run the model using the training/validation data

# Future steps

Its very apparent that 5 epochs is very small. In the future, more epochs would be used on this model. This is called hyperparameter tuning. Learning rate can also be changed as well as the batch size. 

# Known Issues

Limited time and gpu's to train model. 30 epochs on my machine would have been approximately 10 hours. If the gpy "cuda" is unavailable, it can be very difficult to train the model.