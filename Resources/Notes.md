The entire project was guided by o3-mini model, regular free tier as given by OpenAI.  
https://chatgpt.com/share/67b4bc27-6ca0-800b-a587-101e91e2964c

The IPolicyConfig and PolicyConfig interfaces are undocumented by Microsoft and are not officially distributed. However, developers have reverse-engineered these interfaces, and you can find their definitions in various open-source projects. https://chatgpt.com/share/67b4ba5f-97f4-800b-9c6f-bcfcebe8ad54

The o3-mini model was provided with PolicyConfig.h from amate repository as suggested by official Microsoft Vendor 
> Xiaopo Yang - MSFT
> 12,726 Reputation points• Microsoft Vendor
![Screenshot (261)](https://github.com/user-attachments/assets/7734e547-4040-4f6b-8e10-a69eded93c9e)

The official suggestion appeared to be an example repository by https://github.com/amate/

Sources:   
https://learn.microsoft.com/en-us/answers/questions/1162167/how-to-find-audio-devices-that-are-connected  
https://github.com/amate/SetDefaultAudioDevice/blob/540e9b6c55554f27b5a5436505c8e62d6ed253e6/PolicyConfig.h  

Windows SDK Headers:  
https://github.com/hughbe/windows-sdk-headers/blob/89f151d343587fcd5b854a4851fcb8f7187f3ac4/Include/v7.1/endpointvolume.h#L4    
https://github.com/hughbe/windows-sdk-headers/blob/89f151d343587fcd5b854a4851fcb8f7187f3ac4/Include/v7.1/mmdeviceapi.h#L897  

Unused resources:  
https://github.com/ThiefMaster/coreaudio-dotnet/blob/master/CoreAudio/Interfaces/IPolicyConfig.cs  
https://github.com/Belphemur/AudioEndPointLibrary/blob/master/DefSound/PolicyConfig.h  

