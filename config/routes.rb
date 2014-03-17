Thanos::Application.routes.draw do
  root 'dashboard#index'
  post '/upload' => 'upload#create'
end
