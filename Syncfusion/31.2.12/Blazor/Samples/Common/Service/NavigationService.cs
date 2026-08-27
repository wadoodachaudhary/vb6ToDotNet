using Microsoft.AspNetCore.Components;
namespace BlazorDemos.Service {

    public class NavigationService {
        private readonly NavigationManager _navigationManager;

        // Use constructor injection
        public NavigationService(NavigationManager navigationManager){
            _navigationManager = navigationManager;
        }

        // This is a non-static method
        public void GoHome(){
            _navigationManager.NavigateTo("/HF");
        }
    }
}