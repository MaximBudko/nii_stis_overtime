RedmineApp::Application.routes.draw do
    get '/overtime_report', to: 'over_work#generate_overtime_report', as: 'generate_overtime_report'
    post '/overtime_generate', to: 'over_work#do_generate_ov', as: 'do_generate_ov'
end